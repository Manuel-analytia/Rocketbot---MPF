"""
dbtools.py - Decoder/encoder para robot.db de Rocketbot.

Cada fila de la tabla 'bots' tiene 'data' = base64(json(project_dict)).
Project dict tiene:
  - vars: list de {name, data, type, ...}  (variables del modulo del bot)
  - commands: list de comandos, cada uno con
        father: str ('execScriptPython' | 'setVar' | 'evaluateIf' | 'execRocketBotDB' | ...)
        command: str          (Python si execScriptPython; expresion si evaluateIf; valor si setVar)
        description: str
        children: list        (solo si evaluateIf)
        else: list            (solo si evaluateIf)
        execute, disabled, etc.

Uso programatico:
    from dbtools import RobotDB
    db = RobotDB(r'c:\\Repositorios\\MPF - Rocketbot\\Rocketbot---MPF\\robot.db')
    bots = db.list_bots()                    # [(id, name)]
    data = db.read_bot(13)                   # dict completo
    cmd  = db.get_command(13, 0)             # dict del comando
    db.update_command(13, 0, cmd)            # guarda el comando modificado
    db.write_bot(13, data)                   # guarda el bot completo

CLI:
    python dbtools.py list
    python dbtools.py vars 13
    python dbtools.py cmds 32
    python dbtools.py dump 13 0           # script Python del comando a stdout
    python dbtools.py dump 13 0 out.py    # script a archivo
    python dbtools.py grep "RIVERSIDE"    # busca patron en todos los scripts
"""
from __future__ import annotations

import argparse
import base64
import json
import sqlite3
import sys
from pathlib import Path
from typing import Any


DEFAULT_DB = r'c:\Repositorios\MPF - Rocketbot\Rocketbot---MPF\robot.db'


class RobotDB:
    """Wrapper minimal para leer/escribir bots en robot.db."""

    def __init__(self, path: str | Path = DEFAULT_DB):
        self.path = str(path)

    # ---------- low-level ----------
    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self.path)

    @staticmethod
    def _decode(blob: str) -> dict:
        return json.loads(base64.b64decode(blob))

    @staticmethod
    def _encode(data: dict) -> str:
        return base64.b64encode(
            json.dumps(data, ensure_ascii=False).encode('utf-8')
        ).decode('ascii')

    # ---------- listado ----------
    def list_bots(self) -> list[tuple[int, str]]:
        with self._connect() as conn:
            return list(conn.execute('SELECT id, name FROM bots ORDER BY id'))

    # ---------- bot completo ----------
    def read_bot(self, bot_id: int) -> dict:
        with self._connect() as conn:
            row = conn.execute('SELECT data FROM bots WHERE id=?', (bot_id,)).fetchone()
        if not row:
            raise KeyError(f'bot id={bot_id} no existe')
        return self._decode(row[0])

    def write_bot(self, bot_id: int, data: dict) -> None:
        encoded = self._encode(data)
        with self._connect() as conn:
            conn.execute('UPDATE bots SET data=? WHERE id=?', (encoded, bot_id))
            conn.commit()

    # ---------- atajos ----------
    def list_vars(self, bot_id: int) -> list[dict]:
        return self.read_bot(bot_id).get('project', {}).get('vars', [])

    def list_commands(self, bot_id: int) -> list[dict]:
        return self.read_bot(bot_id).get('project', {}).get('commands', [])

    def get_command(self, bot_id: int, cmd_idx: int) -> dict:
        return self.list_commands(bot_id)[cmd_idx]

    def update_command(self, bot_id: int, cmd_idx: int, new_cmd: dict) -> None:
        data = self.read_bot(bot_id)
        data['project']['commands'][cmd_idx] = new_cmd
        self.write_bot(bot_id, data)

    def update_command_script(self, bot_id: int, cmd_idx: int, new_script: str,
                              compile_check: bool = True) -> None:
        """Reemplaza solo el campo .command (codigo Python) preservando el resto."""
        if compile_check:
            compile(new_script, f'<bot{bot_id}_cmd{cmd_idx:02d}>', 'exec')
        data = self.read_bot(bot_id)
        data['project']['commands'][cmd_idx]['command'] = new_script
        self.write_bot(bot_id, data)

    # ---------- helpers de busqueda ----------
    def iter_scripts(self):
        """Itera (bot_id, bot_name, cmd_path, command_dict) para todos los
        execScriptPython, recorriendo recursivamente children/else."""
        for bot_id, name in self.list_bots():
            data = self.read_bot(bot_id)
            for path, cmd in _walk_commands(data['project']['commands']):
                if cmd.get('father') == 'execScriptPython':
                    yield bot_id, name, path, cmd

    def grep(self, pattern: str, ignore_case: bool = False) -> list[tuple[int, str, str, int]]:
        """Busca pattern (substring, no regex) en todos los scripts.
        Retorna (bot_id, bot_name, cmd_path, line_no)."""
        hits = []
        needle = pattern.lower() if ignore_case else pattern
        for bot_id, name, path, cmd in self.iter_scripts():
            script = cmd.get('command', '') or ''
            for ln_no, ln in enumerate(script.split('\n')):
                hay = ln.lower() if ignore_case else ln
                if needle in hay:
                    hits.append((bot_id, name, path, ln_no))
        return hits


def _walk_commands(commands: list[dict], prefix: str = ''):
    """Generator: (path, cmd) para cada comando incluido nested en evaluateIf."""
    for i, cmd in enumerate(commands):
        cur = f'{prefix}cmd_{i:02d}' if not prefix else f'{prefix}.{i}'
        yield cur, cmd
        if cmd.get('children'):
            yield from _walk_commands(cmd['children'], cur + '.children')
        if cmd.get('else'):
            yield from _walk_commands(cmd['else'], cur + '.else')


# ---------------- CLI ----------------
def _cli():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--db', default=DEFAULT_DB)
    sub = ap.add_subparsers(dest='cmd', required=True)

    sub.add_parser('list', help='lista todos los bots (id, name)')

    p_vars = sub.add_parser('vars', help='vars del modulo de un bot')
    p_vars.add_argument('bot_id', type=int)

    p_cmds = sub.add_parser('cmds', help='lista comandos del bot (top-level)')
    p_cmds.add_argument('bot_id', type=int)

    p_dump = sub.add_parser('dump', help='dump script Python del comando')
    p_dump.add_argument('bot_id', type=int)
    p_dump.add_argument('cmd_idx', type=int)
    p_dump.add_argument('out', nargs='?', help='archivo destino (stdout si se omite)')

    p_grep = sub.add_parser('grep', help='busca substring en todos los scripts')
    p_grep.add_argument('pattern')
    p_grep.add_argument('-i', '--ignore-case', action='store_true')

    args = ap.parse_args()
    db = RobotDB(args.db)

    if args.cmd == 'list':
        for bid, name in db.list_bots():
            print(f'  {bid:3d}  {name}')
    elif args.cmd == 'vars':
        for v in db.list_vars(args.bot_id):
            print(f"  {v.get('name','?'):40s} = {v.get('data','')!r}")
    elif args.cmd == 'cmds':
        for i, c in enumerate(db.list_commands(args.bot_id)):
            dis = ' [DISABLED]' if c.get('disabled') else ''
            tag = c.get('father', '?')
            desc = (c.get('description', '') or '')[:60]
            print(f"  CMD{i:02d} [{tag}] {desc}{dis}")
    elif args.cmd == 'dump':
        cmd = db.get_command(args.bot_id, args.cmd_idx)
        if cmd.get('father') != 'execScriptPython':
            print(f'# WARN no es execScriptPython, father={cmd.get("father")}', file=sys.stderr)
        script = cmd.get('command', '')
        if args.out:
            Path(args.out).write_text(script, encoding='utf-8')
            print(f'-> {args.out}', file=sys.stderr)
        else:
            print(script)
    elif args.cmd == 'grep':
        hits = db.grep(args.pattern, ignore_case=args.ignore_case)
        for bid, name, path, ln in hits:
            print(f'  bot {bid} ({name}) {path} L{ln}')


if __name__ == '__main__':
    _cli()
