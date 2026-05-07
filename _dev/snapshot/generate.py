"""
Genera snapshot del estado actual de cada bot en _dev/snapshot/bot_NN/
Uso:
    python _dev/snapshot/generate.py

Resultado por bot:
    _dev/snapshot/bot_NN/meta.json    - vars del modulo + arbol de comandos sin scripts
    _dev/snapshot/bot_NN/cmd_XX.py    - scripts Python (un archivo por execScriptPython)

Para nested (children/else de evaluateIf) los scripts van como
    cmd_07.children.04.py
para que el path sea legible y grepeable.
"""
from __future__ import annotations
import json
import sys
from pathlib import Path

# permitir importar _dev/lib/dbtools.py
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'lib'))
from dbtools import RobotDB, _walk_commands  # noqa: E402


HERE = Path(__file__).resolve().parent


def snapshot_bot(db: RobotDB, bot_id: int, name: str) -> int:
    """Escribe meta.json + cmd_XX.py para un bot. Retorna # scripts escritos."""
    out_dir = HERE / f'bot_{bot_id}_{name}'
    out_dir.mkdir(parents=True, exist_ok=True)

    # Limpia scripts viejos (no archivos meta)
    for old in out_dir.glob('cmd_*.py'):
        old.unlink()

    data = db.read_bot(bot_id)
    project = data.get('project', {})

    # Recolecta scripts y reemplaza el campo 'command' por placeholder en el meta
    n_scripts = 0
    meta_commands: list = []
    for path, cmd in _walk_commands(project.get('commands', [])):
        # generamos meta con shape original recursivo abajo
        pass

    def _strip_scripts(commands: list, prefix: str = '') -> list:
        nonlocal n_scripts
        out = []
        for i, cmd in enumerate(commands):
            cur = f'{prefix}cmd_{i:02d}' if not prefix else f'{prefix}.{i:02d}'
            cp = {k: v for k, v in cmd.items() if k not in ('children', 'else')}
            if cmd.get('father') == 'execScriptPython':
                fname = cur.replace('.', '_') + '.py'
                (out_dir / fname).write_text(cmd.get('command', '') or '', encoding='utf-8')
                cp['command'] = f'__FILE__:{fname}'
                n_scripts += 1
            if cmd.get('children'):
                cp['children'] = _strip_scripts(cmd['children'], cur + '.children.')
            if cmd.get('else'):
                cp['else'] = _strip_scripts(cmd['else'], cur + '.else.')
            out.append(cp)
        return out

    meta = {
        'bot_id': bot_id,
        'name': name,
        'vars': project.get('vars', []),
        'commands': _strip_scripts(project.get('commands', [])),
    }
    (out_dir / 'meta.json').write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding='utf-8'
    )
    return n_scripts


def main():
    db = RobotDB()
    total_scripts = 0
    print(f'Snapshot a {HERE}')
    for bot_id, name in db.list_bots():
        n = snapshot_bot(db, bot_id, name)
        print(f'  bot {bot_id:3d} ({name}): {n} scripts')
        total_scripts += n
    print(f'Total: {total_scripts} scripts escritos.')


if __name__ == '__main__':
    main()
