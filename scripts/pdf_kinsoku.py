"""ReportLab の日本語折り返しで、行頭禁則の文字が2つ続く場合も前の行へぶら下げる補正。

ReportLab 4.x の cjkFragSplit は、はみ出した1文字が行頭禁則（「）」「。」など）なら前の行へ
ぶら下げるが、2文字目は扱わない。そのため「…（同 項目15）。」の「。」が次の行の先頭に来る。
この補正は、ぶら下げた直後の文字も行頭禁則なら、もう1文字だけ同じ行へぶら下げる。
"""
import inspect

from reportlab.lib import textsplit as _textsplit
from reportlab.platypus import paragraph as _paragraph

_ORIGINAL_HANG = '''                if u not in ALL_CANNOT_START and i>lineStartPos+1:
                    #otherwise we need to push the character back
                    #the i>lineStart+1 condition ensures progress
                    i -= 1
                    extraSpace += w
'''
_PATCHED_HANG = _ORIGINAL_HANG + '''                elif u in ALL_CANNOT_START and i<nU and U[i] in ALL_CANNOT_START:
                    #行頭禁則が2つ続く場合（「）。」など）は2文字目も同じ行へぶら下げる
                    w2 = U[i].width
                    if hasattr(w2,'normalizedValue'):
                        w2._normalizer = maxWidth
                        w2 = w2.normalizedValue(maxWidth)
                    widthUsed += w2
                    extraSpace -= w2
                    i += 1
'''
# 1つの書体だけの段落は textsplit.dumbSplit が折り返す（変数名が c・word・widths）
_ORIGINAL_HANG_DUMB = '''            if c not in ALL_CANNOT_START and i>lineStartPos+1:
                #otherwise we need to push the character back
                #the i>lineStart+1 condition ensures progress
                i -= 1
                extraSpace += w
'''
_PATCHED_HANG_DUMB = _ORIGINAL_HANG_DUMB + '''            elif c in ALL_CANNOT_START and i<len(word) and word[i] in ALL_CANNOT_START:
                #行頭禁則が2つ続く場合（「）。」など）は2文字目も同じ行へぶら下げる
                widthUsed += widths[i]
                extraSpace -= widths[i]
                i += 1
'''
_applied = False


class KinsokuPatchError(RuntimeError):
    """ReportLab の版が想定と違い、補正を当てられないとき。"""


def _patched(module, name, original, patched):
    source = inspect.getsource(getattr(module, name))
    if source.count(original) != 1:
        raise KinsokuPatchError(f'ReportLab の {name} が想定と違うため、行頭禁則の補正を当てられません')
    namespace = dict(vars(module))
    exec(compile(source.replace(original, patched), module.__file__, 'exec'), namespace)
    return namespace[name]


def apply():
    """複数書体用（paragraph.cjkFragSplit）と1書体用（textsplit.dumbSplit）を補正版に差し替える。"""
    global _applied
    if _applied:
        return
    cjk = _patched(_paragraph, 'cjkFragSplit', _ORIGINAL_HANG, _PATCHED_HANG)
    dumb = _patched(_textsplit, 'dumbSplit', _ORIGINAL_HANG_DUMB, _PATCHED_HANG_DUMB)
    _paragraph.cjkFragSplit = cjk
    _textsplit.dumbSplit = dumb
    _applied = True
