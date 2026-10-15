"""UTF-8/CP932を厳密に復号し、自動で選べない場合は利用者へ返す。"""

import codecs
from dataclasses import dataclass


@dataclass(frozen=True)
class EncodingDecision:
    candidates: tuple[str, ...]
    selected: str | None
    previews: dict[str, str]
    reason: str


class AmbiguousEncodingError(ValueError):
    pass


class EncodingSelectionCancelled(Exception):
    pass


def detect_encoding(raw: bytes) -> EncodingDecision:
    if raw.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE, codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        raise ValueError('UTF-16/UTF-32のCSVは未対応です。UTF-8またはCP932へ変換してください。')
    if b'\x00' in raw:
        raise ValueError('NUL文字を含むファイルは自動判定できません。UTF-16またはバイナリ形式の可能性があります。')
    if raw.startswith(codecs.BOM_UTF8):
        text = raw[len(codecs.BOM_UTF8):].decode('utf-8', errors='strict')
        return EncodingDecision(('utf-8',), 'utf-8', {'utf-8': text[:1200]}, 'UTF-8 BOM')
    decoded = {}
    for encoding in ('utf-8', 'cp932'):
        try:
            decoded[encoding] = raw.decode(encoding, errors='strict')
        except UnicodeDecodeError:
            pass
    if not decoded:
        raise ValueError('UTF-8・CP932のどちらでも読み込めません。文字コードを確認してください。')
    candidates = tuple(decoded)
    previews = {key: value[:1200] for key, value in decoded.items()}
    if len(decoded) == 1:
        return EncodingDecision(candidates, candidates[0], previews, '復号できる候補が1つ')
    if decoded['utf-8'] == decoded['cp932']:
        return EncodingDecision(candidates, 'utf-8', previews, '候補の内容が同じためUTF-8を使用')
    return EncodingDecision(candidates, None, previews, '両方で読めるため選択が必要')


def resolve_encoding(raw, choose=None):
    decision = detect_encoding(raw)
    if decision.selected is not None:
        return decision.selected
    if choose is None:
        raise AmbiguousEncodingError('UTF-8とCP932のどちらで開くか選択してください。')
    selected = choose(decision)
    if selected is None:
        raise EncodingSelectionCancelled()
    if selected not in decision.candidates:
        raise ValueError('判定候補にない文字コードです。')
    return selected
