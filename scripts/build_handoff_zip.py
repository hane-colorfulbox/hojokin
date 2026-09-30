"""（廃止）引き継ぎパックの ZIP 配布（hojokin-handoff.zip）。

2026-09-30 に、補助金業務の資料とスキルは組織のプラグイン（hojokin）で配る方式へ移行した。
配布物は scripts/build_plugin.py で組み立てる。許可リストの正本は scripts/distribution_manifest.py。
旧パックの利用者を移行させる最終版 ZIP は `python scripts/build_plugin.py --retire-zip <出力先>` で作る。
"""
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

MESSAGE = ('build_handoff_zip.py は廃止しました（組織プラグインでの配布に移行）。\n'
           '  配布物: python scripts/build_plugin.py --out <プラグイン用リポジトリ>\n'
           '  旧パックの退役版: python scripts/build_plugin.py --retire-zip _dist/hojokin-handoff.zip')


def main() -> int:
    print(MESSAGE)
    return 2


if __name__ == '__main__':
    raise SystemExit(main())
