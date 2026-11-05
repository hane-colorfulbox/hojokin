"""pytest の収集設定。

tests/plugin は配布版プラグインを組み立ててから実行するスクリプト群で、
`python -B tests/plugin/run_all.py` から起動する（環境変数で配布版の場所を受け取る）。
pytest で直接集めると SystemExit で INTERNALERROR になるため、収集対象から外す。
"""
collect_ignore = ['plugin']
