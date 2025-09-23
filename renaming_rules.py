"""
Движок правил распознавания и форматирования имен.

Цели:
- Надёжное распознавание семейства (папки назначения).
- Нормализация ОС-имени (fs_name) строго по правилам коллеги.
- Отдельная нормализация "display name" для Excel (например, AMSXXXX_X -> AMSXXXX/X).

Ключевые семейства с особыми правилами до наполнения библиотеки:
DAN, DIN, HST, ISO (включая ISO IEC / ISO TR / ISO TS / ISO SAE(PAS)), LN, MEP, MS, NAS, NE, TAN,
AIPI, AN, ARP, AMS (+QPL), D8, D6, D2, а также BS EN ISO.
"""

import re
import yaml
from typing import Any, Dict, List, Optional, Tuple
from pathlib import Path

class FileNameFormatter:
    """
    Форматирование имени файла.

    Правила:
    - format_fs_name(): имя ОС-файла (всегда с расширением).
    - format_display_name(): заголовок для Excel (может отличаться, напр. AMSxxxx_x -> AMSxxxx/x).
    """

    ABBREVS_WITH_SPACE = {'MEP', 'NE', 'SPM', 'TAN'}  # MEP XX-XXX; NE XX-XXX; ...
    ABBREVS_WITHOUT_SPACE = {
        'AIPI','AN','ARP','BAC','BACB','BACC','BACD','BACF','BACI','BACJ',
        'BACL','BACN','BACP','BACR','BACS','BACT','BACV','BACW',
        'DAN','LN','MS','NAS','HST'
    }

    ABBREVS_TO_UPPERCASE = {
        'HST','MS','NAS','СП','MIL','ASTM','QQ','LN','BS',
        'BAC','BACB','BACC','BACD','BACF','BACI','BACJ',
        'BACL','BACN','BACP','BACR','BACS','BACT','BACV','BACW'
    }

    ACRONYMS_TREATED_AS_BASE = {'AMS', 'ASTM', 'ASME', 'FED', 'MIL', 'QQ', 'TN', 'ASD', 'AWS', 'PDS'}
    DIN_COMPOUND = ('EN', 'ISO', 'SAE', 'SPEC')

    def __init__(self, force_uppercase: bool = True):
        self.force_uppercase = force_uppercase

    def _split(self, basename: str) -> Tuple[str, str]:
        p = Path(basename)
        return p.stem, p.suffix  # suffix включает точку или пусто

    def _squash_spaces(self, s: str) -> str:
        return re.sub(r'\s+', ' ', s).strip()

    def _normalize_aipi(self, name: str) -> str:
        """AIPIХХ-ХХ-ХХХ без пробела после префикса."""
        m = re.match(r'^(?i:AIPI)\D*(\d{2})\D*(\d{2})\D*(\d{3,})', name)
        if m:
            return f"AIPI{m.group(1)}-{m.group(2)}-{m.group(3)}"
        return re.sub(r'^(?i:AIPI)\s+', 'AIPI', name)

    def _normalize_qpl_name(self, name: str) -> str:
        """
        Приводит QPL-имена к шаблону «<BASE> QPL [REV ... | ...]».
        - отделяет QPL от базового обозначения (BMS8-79_QPL -> BMS8-79 QPL)
        - для файлов «QPL BMS…» переставляет QPL в конец
        - нормализует ревизии: «_D» -> «REV D», «REV.» -> «REV» и т.п.
        - удаляет лишние маркеры вроде «TO»/«SAE», которые встречаются в исходниках
        """

        if 'QPL' not in name.upper():
            return name

        working = name.replace('_', ' ')
        working = re.sub(r'(?i)REV[\.:]', 'REV ', working)
        working = re.sub(r'(?i)ISSUE[\.:]', 'ISSUE ', working)
        working = re.sub(r'(?i)QPL\s*(?:TO|TO:)', 'QPL TO ', working)
        working = re.sub(r'(?i)([A-Z0-9])[\s._-]*QPL', r'\1 QPL', working)
        working = re.sub(r'(?i)QPL[\s._-]*([A-Z0-9])', r'QPL \1', working)
        tokens_raw = re.split(r'\s+', working.strip())
        tokens: List[str] = []
        for tok in tokens_raw:
            if not tok:
                continue
            upper_tok = tok.upper()
            if upper_tok.startswith('TO') and len(tok) > 2:
                remainder = tok[2:]
                remainder = remainder.lstrip(':-._')
                tokens.append(tok[:2])
                if remainder:
                    tokens.append(remainder)
                continue
            tokens.append(tok)
        if not tokens:
            return name

        skip_words = {'QPL', 'TO', 'TO:', 'SAE', 'SPEC', 'STANDARD', 'STD'}
        base_idx = -1
        for idx, tok in enumerate(tokens):
            upper_tok = tok.upper().strip('.')
            if upper_tok in skip_words:
                continue
            if any(ch.isdigit() for ch in tok):
                base_idx = idx
                break
        if base_idx == -1:
            for idx, tok in enumerate(tokens):
                if tok.upper() not in skip_words:
                    base_idx = idx
                    break
        if base_idx == -1:
            base_idx = 0

        base_token = tokens[base_idx]
        rest_tokens = [tok for idx, tok in enumerate(tokens) if idx != base_idx and tok.upper() not in skip_words]

        cleaned_rest = []
        keyword_map = {
            'REVISION': 'REV',
            'REV.': 'REV',
            'REVISION:': 'REV',
            'REV:': 'REV',
        }
        revision_keywords = {'REV', 'ISSUE', 'ED', 'EDITION', 'CHG', 'CHANGE', 'AMDT', 'AMENDMENT', 'MOD'}
        for token in rest_tokens:
            token = token.strip()
            if not token:
                continue
            stripped = token.rstrip('.,;:')
            mapped = keyword_map.get(stripped.upper(), stripped)
            upper = mapped.upper()
            if upper.startswith('ISSUE') and len(upper) > len('ISSUE'):
                suffix = mapped[len('ISSUE'):].strip()
                if suffix and re.fullmatch(r'[A-Z0-9]{1,3}', suffix.upper()):
                    cleaned_rest.append('ISSUE')
                    cleaned_rest.append(suffix)
                    continue
            if upper.startswith('REV') and len(upper) > len('REV') and upper not in revision_keywords:
                suffix = mapped[len('REV'):].strip('-_. ')
                if suffix and re.fullmatch(r'[A-Z0-9]{1,3}', suffix.upper()):
                    cleaned_rest.append('REV')
                    cleaned_rest.append(suffix)
                    continue
            if mapped:
                cleaned_rest.append(mapped)

        if cleaned_rest:
            first = cleaned_rest[0].upper()
            if first not in revision_keywords and re.fullmatch(r'[A-Z]{1,3}', first):
                cleaned_rest.insert(0, 'REV')

        ordered = [base_token] + (['QPL'] if 'QPL' not in base_token.upper() else []) + cleaned_rest
        normalized = ' '.join(ordered)
        return normalized

    def _format_bms_ams_qpl(self, name: str) -> str:
        if 'QPL' not in name.upper():
            return name

        # Разворачиваем все вхождения ".QPL" и " QPL" в единый маркер
        collapsed = re.sub(r'(?i)\.?\s*QPL\b', ' QPL ', name)
        collapsed = re.sub(r'\s+', ' ', collapsed).strip()
        tokens = collapsed.split(' ')
        if not tokens:
            return name

        qpl_idx = next((i for i, tok in enumerate(tokens) if tok.upper() == 'QPL'), -1)
        if qpl_idx == -1:
            return name

        base_tokens = tokens[:qpl_idx]
        suffix_tokens = tokens[qpl_idx + 1:]

        base = ' '.join(base_tokens).strip()
        revision = ' '.join(suffix_tokens).strip()

        base = re.sub(r'(?i)\.?\s*QPL\b', '', base).strip()
        revision = re.sub(r'(?i)\.?\s*QPL\b', '', revision).strip()

        combined = ' '.join([part for part in (base, revision) if part])
        combined = combined.strip(' .')
        if not combined:
            combined = base or revision or ''

        if not combined:
            return 'QPL'

        combined = re.sub(r'\s+', ' ', combined).strip()
        formatted = combined.rstrip('. ')
        formatted = re.sub(r'(?i)(?:\.\s*QPL)+$', '', formatted).strip()
        if formatted.upper().endswith('.QPL'):
            formatted = formatted[:-4].rstrip()

        return f"{formatted}.QPL"

    def _ensure_tan_spacing(self, name: str) -> str:
        """
        Для TAN: допускаем только шаблон «TAN XXX XX» без суффиксов/префиксов.
        """
        m = re.search(r'(?i:TAN)\D*(\d{3})\D*(\d{2})', name)
        if m:
            return f"TAN {m.group(1)} {m.group(2)}"
        return name

    def format_fs_name(self, basename: str, folder: str, original_abbrev: str) -> str:
        """
        Вернуть нормализованное имя ОС-файла с расширением. Расширение источника сохраняется.
        """
        stem, ext = self._split(basename)
        ext = ext or ""  # страховка сохранения исходного расширения
        folder_up = folder.upper()
        orig_up = original_abbrev.upper()
        name = stem

        # Убираем префикс "SAE " при наличии
        name = re.sub(r'^(?i:SAE)[\s_-]*', '', name).strip()

        # AIPI / AN / ARP: без пробела после префикса
        if folder_up == 'AIPI':
            name = self._normalize_aipi(name)
        elif folder_up in ('AN', 'ARP'):
            name = re.sub(rf'^(?i:{folder_up})\s+', folder_up, name)

        # BAC, DAN, LN, MS, NAS, HST: без пробела
        if folder_up in self.ABBREVS_WITHOUT_SPACE:
            name = re.sub(rf'^(?i:{folder_up})[\s_-]+', folder_up, name)

        # MEP/NE/SPM/TAN: обязательный пробел
        if folder_up in self.ABBREVS_WITH_SPACE:
            name = re.sub(rf'^(?i:{folder_up})[\s_-]*', f'{folder_up} ', name)

        # TAN — доп. пробел после трёх цифр
        if folder_up == 'TAN':
            name = self._ensure_tan_spacing(name)

        # DIN-комбинации
        if folder_up == 'DIN':
            if re.match(r'^(?i:DIN)\s*(?:_|-|\s)*EN(?:\s*(?:_|-|\s)*ISO)?', name):
                name = re.sub(r'^(?i:DIN)\s*(?:_|-|\s)*EN', 'DIN EN', name).strip()
                name = re.sub(r'^(?i:DIN EN)\s*(?:_|-|\s)*ISO', 'DIN EN ISO', name).strip()
            elif re.match(r'^(?i:DIN)\s*(?:_|-|\s)*ISO', name):
                name = re.sub(r'^(?i:DIN)\s*(?:_|-|\s)*ISO', 'DIN ISO', name).strip()
            elif re.match(r'^(?i:DIN)\s*(?:_|-|\s)*SAE(?:\s*(?:_|-|\s)*SPEC)?', name):
                name = re.sub(r'^(?i:DIN)\s*(?:_|-|\s)*SAE(\s*(?:_|-|\s)*SPEC)?', r'DIN SAE\1', name).strip()
            else:
                name = re.sub(r'^(?i:DIN)[\s_-]+', 'DIN ', name).strip()

        # ISO комбинации: IEC, SAE (PAS), TR, TS
        if folder_up == 'ISO':
            s = name
            # ISO IEC
            if re.match(r'^ISO\s*(?:_|-|\.)*\s*IEC', s, flags=re.IGNORECASE):
                tail = re.sub(r'^ISO\s*(?:_|-|\.)*\s*IEC', '', s, flags=re.IGNORECASE).lstrip(' _-.')
                name = f"ISO IEC {tail}".strip()
            # ISO SAE (PAS)
            elif re.match(r'^ISO\s*(?:_|-|\.)*\s*SAE(?:\s*(?:_|-|\.)*\s*PAS)?', s, flags=re.IGNORECASE):
                tail = re.sub(r'^ISO\s*(?:_|-|\.)*\s*SAE(?:\s*(?:_|-|\.)*\s*PAS)?', '', s, flags=re.IGNORECASE).lstrip(' _-.')
                if re.match(r'^(?i:PAS)\b', tail):
                    tail = re.sub(r'^(?i:PAS)\s*', '', tail).lstrip(' _-.')
                    name = f"ISO PAS {tail}".strip()
                else:
                    name = f"ISO {tail}".strip()
            # ISO TR / ISO TS
            elif re.match(r'^ISO\s*(?:_|-|\.)*\s*TR', s, flags=re.IGNORECASE):
                tail = re.sub(r'^ISO\s*(?:_|-|\.)*\s*TR', '', s, flags=re.IGNORECASE).lstrip(' _-.')
                name = f"ISO TR {tail}".strip()
            elif re.match(r'^ISO\s*(?:_|-|\.)*\s*TS', s, flags=re.IGNORECASE):
                tail = re.sub(r'^ISO\s*(?:_|-|\.)*\s*TS', '', s, flags=re.IGNORECASE).lstrip(' _-.')
                name = f"ISO TS {tail}".strip()
            else:
                # Простой ISO
                name = re.sub(r'^ISO(?:\s*|_|-|\.)+', 'ISO ', s, flags=re.IGNORECASE)
                name = re.sub(r'^ISO(?=\d)', 'ISO ', name, flags=re.IGNORECASE)

        # ASTM — дефис после ASTM-
        if folder_up == 'ASTM':
            if not re.match(r'^(?i:ASTM)-', name):
                name = re.sub(r'^(?i:ASTM)([A-Z])', r'ASTM-\1', name)

        # MIL — MIL-STD/-PRF/-DTL/-HDBK/-SPEC
        if folder_up == 'MIL':
            m = re.match(r'^(?i:MIL)([A-Z]+)', name)
            if m:
                suffix = m.group(1).upper()
                if suffix in ['STD', 'PRF', 'DTL', 'HDBK', 'SPEC']:
                    name = re.sub(r'^(?i:MIL)'+suffix, f'MIL-{suffix}', name)

        # PPP — PPP-H-
        if folder_up == 'PPP':
            name = re.sub(r'^(?i:PPPH)-', 'PPP-H-', name)

        # AMS — убрать пробелы после AMS; AMS-<SUFFIX> при наличии буквенного суффикса
        if folder_up == 'AMS':
            name = re.sub(r'^(?i:AMS)\s+', 'AMS', name)
            m = re.match(r'^(?i:AMS)([A-Z]+)', name)
            if m:
                suffix = m.group(1).upper()
                name = re.sub(r'^(?i:AMS)'+suffix, f'AMS-{suffix}', name)

        # BS EN ISO — британские стандарты
        if folder_up == 'BS':
            if re.match(r'^(?i:BS)\s*(?:_|-|\.)*\s*EN(?:\s*(?:_|-|\.)*\s*ISO)?', name):
                name = re.sub(r'^(?i:BS)\s*(?:_|-|\.)*\s*EN', 'BS EN', name)
                name = re.sub(r'^(?i:BS EN)\s*(?:_|-|\.)*\s*ISO', 'BS EN ISO', name)
            else:
                name = re.sub(r'^(?i:BS)[\s_-]+', 'BS ', name)

        # --- QPL: особая логика ---
        name = self._normalize_qpl_name(name)

        if folder_up in {'BMS', 'AMS'}:
            name = self._format_bms_ams_qpl(name)

        # Гарантируем верхний регистр ряда префиксов
        for pref in self.ABBREVS_TO_UPPERCASE:
            name = re.sub(rf'^(?i:{pref})', pref, name)

        name = self._squash_spaces(name)
        if self.force_uppercase:
            name = name.upper()
        return name + ext

    def format_display_name(self, fs_basename: str, folder: str, original_abbrev: str) -> str:
        """
        Представление для Excel (может отличаться от OS-имени).
        Примеры:
        - AMSXXXX_X -> AMSXXXX/X
        - AMS-QQ-P-416_5 -> AMS-QQ-P-416/5
        """
        stem, _ext = self._split(fs_basename)
        display = stem
        if folder.upper() == 'AMS':
            if '_' in stem:
                parts = stem.split('_')
                if len(parts[-1]) <= 2:
                    display = '/'.join(['_'.join(parts[:-1]), parts[-1]])
        display = self._squash_spaces(display)
        return display.upper()  # Excel: требование UPPERCASE


class RulesEngine:
    """
    Движок распознавания семейства по имени файла (регулярки + пост-валидация).
    """

    def __init__(self, rules_file: Optional[str] = None, logger=None):
        self.logger = logger
        self.detect_rules: List[Dict[str, Any]] = []
        self._load_yaml_rules(rules_file)

    DEFAULT_RULES_YAML = """
detect:
  # QPL to TARGET
  - name: QPL
    priority: 300
    pattern: '^(?:QPL)(?:[\\s_-]*to)?[\\s_-]*([A-Z]{2,4})[\\s_-]*(.+)$'
    folder_from_group: 1
    original_abbrev: 'QPL'

  # BAC* subfamilies (должны идти раньше общего BAC!)
  - name: BACB
    priority: 189
    pattern: '^BACB[\\s\\-A-Z0-9_]+' 
    folder: 'BACB'
    original_abbrev: 'BACB'
  - name: BACC
    priority: 189
    pattern: '^BACC[\\s\\-A-Z0-9_]+' 
    folder: 'BACC'
    original_abbrev: 'BACC'
  - name: BACD
    priority: 189
    pattern: '^BACD[\\s\\-A-Z0-9_]+' 
    folder: 'BACD'
    original_abbrev: 'BACD'
  - name: BACF
    priority: 189
    pattern: '^BACF[\\s\\-A-Z0-9_]+' 
    folder: 'BACF'
    original_abbrev: 'BACF'
  - name: BACI
    priority: 189
    pattern: '^BACI[\\s\\-A-Z0-9_]+' 
    folder: 'BACI'
    original_abbrev: 'BACI'
  - name: BACJ
    priority: 189
    pattern: '^BACJ[\\s\\-A-Z0-9_]+' 
    folder: 'BACJ'
    original_abbrev: 'BACJ'
  - name: BACL
    priority: 189
    pattern: '^BACL[\\s\\-A-Z0-9_]+' 
    folder: 'BACL'
    original_abbrev: 'BACL'
  - name: BACN
    priority: 189
    pattern: '^BACN[\\s\\-A-Z0-9_]+' 
    folder: 'BACN'
    original_abbrev: 'BACN'
  - name: BACP
    priority: 189
    pattern: '^BACP[\\s\\-A-Z0-9_]+' 
    folder: 'BACP'
    original_abbrev: 'BACP'
  - name: BACR
    priority: 189
    pattern: '^BACR[\\s\\-A-Z0-9_]+' 
    folder: 'BACR'
    original_abbrev: 'BACR'
  - name: BACS
    priority: 189
    pattern: '^BACS[\\s\\-A-Z0-9_]+' 
    folder: 'BACS'
    original_abbrev: 'BACS'
  - name: BACT
    priority: 189
    pattern: '^BACT[\\s\\-A-Z0-9_]+' 
    folder: 'BACT'
    original_abbrev: 'BACT'
  - name: BACV
    priority: 189
    pattern: '^BACV[\\s\\-A-Z0-9_]+' 
    folder: 'BACV'
    original_abbrev: 'BACV'
  - name: BACW
    priority: 189
    pattern: '^BACW[\\s\\-A-Z0-9_]+' 
    folder: 'BACW'
    original_abbrev: 'BACW'

  # AIPI (без пробела)
  - name: AIPI
    priority: 190
    pattern: '^AIPI\\s*\\d'
    folder: 'AIPI'
    original_abbrev: 'AIPI'

  # AN / ARP — без пробела
  - name: AN
    priority: 188
    pattern: '^AN\\s*\\d'
    folder: 'AN'
    original_abbrev: 'AN'
  - name: ARP
    priority: 188
    pattern: '^ARP\\s*\\d'
    folder: 'ARP'
    original_abbrev: 'ARP'

  # AMS (все формы)
  - name: AMS
    priority: 180
    pattern: '^(?:SAE[\\s_-]*)?AMS[-\\sA-Z0-9_]+' 
    folder: 'AMS'
    original_abbrev: 'AMS'

  # D8 / D6 / D2 (только с дефисом)
  - name: D8
    priority: 170
    pattern: '^D8-\\w+'
    folder: 'D8'
    original_abbrev: 'D8'
  - name: D6
    priority: 170
    pattern: '^D6-\\w+'
    folder: 'D6'
    original_abbrev: 'D6'
  - name: D2
    priority: 170
    pattern: '^D2-\\w+'
    folder: 'D2'
    original_abbrev: 'D2'

  # DIN combos
  - name: DIN EN ISO
    priority: 160
    pattern: '^DIN(?:\\s*|_|-)*EN(?:\\s*|_|-)*ISO'
    folder: 'DIN'
    original_abbrev: 'DIN EN ISO'
  - name: DIN EN
    priority: 150
    pattern: '^DIN(?:\\s*|_|-)*EN'
    folder: 'DIN'
    original_abbrev: 'DIN EN'
  - name: DIN ISO
    priority: 150
    pattern: '^DIN(?:\\s*|_|-)*ISO'
    folder: 'DIN'
    original_abbrev: 'DIN ISO'
  - name: DIN SAE
    priority: 150
    pattern: '^DIN(?:\\s*|_|-)*SAE(?:\\s*SPEC)?'
    folder: 'DIN'
    original_abbrev: 'DIN SAE'

  # ISO IEC / ISO SAE / ISO TS / ISO TR / ISO
  - name: ISO IEC
    priority: 140
    pattern: '^ISO(?:\\s*|_|-)*IEC'
    folder: 'ISO'
    original_abbrev: 'ISO IEC'
  - name: ISO SAE (PAS)
    priority: 140
    pattern: '^ISO(?:\\s*|_|-)*SAE(?:\\s*PAS)?'
    folder: 'ISO'
    original_abbrev: 'ISO SAE'
  - name: ISO TS
    priority: 130
    pattern: '^ISO(?:\\s*|_|-)*TS'
    folder: 'ISO'
    original_abbrev: 'ISO TS'
  - name: ISO TR
    priority: 130
    pattern: '^ISO(?:\\s*|_|-)*TR'
    folder: 'ISO'
    original_abbrev: 'ISO TR'
  - name: ISO
    priority: 110
    pattern: '^ISO(?:\\s*|_|-|\\.)*\\d'
    folder: 'ISO'
    original_abbrev: 'ISO'

  # BS EN ISO / BS EN / BS
  - name: BS EN ISO
    priority: 105
    pattern: '^BS(?:\\s*|_|-)*EN(?:\\s*|_|-)*ISO'
    folder: 'BS'
    original_abbrev: 'BS EN ISO'
  - name: BS EN
    priority: 100
    pattern: '^BS(?:\\s*|_|-)*EN'
    folder: 'BS'
    original_abbrev: 'BS EN'
  - name: BS
    priority: 90
    pattern: '^BS\\s*\\d'
    folder: 'BS'
    original_abbrev: 'BS'

  # Общий BAC — в самом конце этой группы
  - name: BAC generic
    priority: 80
    pattern: '^BAC[\\s\\-A-Z0-9_]+'
    folder: 'BAC'
    original_abbrev: 'BAC'
"""

    def _load_yaml_rules(self, rules_file: Optional[str]):
        raw = self.DEFAULT_RULES_YAML
        if rules_file:
            try:
                with open(rules_file, 'r', encoding='utf-8') as f:
                    raw = f.read()
            except Exception:
                if self.logger:
                    self.logger.warning(f"Не удалось открыть rules.yaml: {rules_file}, используются дефолтные.")
        try:
            data = yaml.safe_load(raw) or {}
            self.detect_rules = data.get('detect', [])
            self.detect_rules.sort(key=lambda r: r.get('priority', 0), reverse=True)
        except Exception as e:
            self.detect_rules = []
            if self.logger:
                self.logger.warning(f"Ошибка загрузки YAML правил: {e}. Правила отключены — будет использоваться базовый поиск.")

    def apply_rules(self, filename: str) -> Optional[Tuple[str, str]]:
        for rule in self.detect_rules:
            pat = rule.get('pattern')
            if not pat:
                continue
            m = re.search(pat, filename, flags=re.IGNORECASE)
            if not m:
                continue

            folder: Optional[str] = rule.get('folder')
            original: Optional[str] = rule.get('original_abbrev')

            if rule.get('folder_from_group'):
                gi = int(rule['folder_from_group'])
                folder = m.group(gi).upper()

            if rule.get('original_abbrev_from_group'):
                gi = int(rule['original_abbrev_from_group'])
                original = m.group(gi)

            if not folder:
                continue
            if not original:
                original = folder

            if self.is_valid_standard(filename, folder, original):
                return folder, original
        return None

    def is_valid_standard(self, filename: str, folder: str, original_abbrev: str) -> bool:
        """
        Фильтрация «мусора» по семействам из комментариев коллег.
        """
        name_lower = filename.lower()
        folder_up = folder.upper()

        # Русские исключения
        if folder == 'ПИ' and any(w in name_lower for w in ['письмо', 'письма']):
            return False
        if folder == 'СП' and any(w in name_lower for w in ['список', 'справка']):
            return False

        # AMS — отсечь Airbus Illustrated Parts и AIPI/AIPC в AMS
        if folder_up == 'AMS':
            bad_tokens = ['illustrated', 'aipi', 'aipc']
            if any(t in name_lower for t in bad_tokens):
                return False

        # AN: после AN обязательно цифры
        if folder_up == 'AN':
            if not re.search(r'^(?i:AN)\s*\d', filename):
                return False

        # AS — только AS<цифры> или AS-<цифры>; отсекаем чертежи
        if folder_up == 'AS':
            if not re.match(r'^(?i:AS)\s*-?\d', filename):
                return False
            if 'dwg' in name_lower or 'drawing' in name_lower:
                return False

        # ASN — отсекаем явные нерелевантные
        if folder_up == 'ASN':
            if any(b in name_lower for b in ['wiring diagrams', 'entertainment']):
                return False

        # D8/D6/D2 — строгие
        if folder_up == 'D8':
            if not re.match(r'^(?i:D8)-\w+', filename):
                return False
        if folder_up == 'D6':
            if not re.match(r'^(?i:D6)-\w+', filename):
                return False
            if 'manual' in name_lower:
                return False
        if folder_up == 'D2':
            if not re.match(r'^(?i:D2)-\w+', filename):
                return False

        # DIN — EN/ISO/SAE(SPEC) или DIN <номер>
        if folder_up == 'DIN':
            ok = (re.match(r'^(?i:DIN)\s+(EN|ISO|SAE)(?:\s+SPEC)?\s+\S', filename)
                  or re.match(r'^(?i:DIN)\s+\d', filename))
            if not ok:
                return False

        # ISO — допускаем разделители '_', '-', '.'
        if folder_up == 'ISO':
            sep = r'(?:\s*|_|-|\.)*'
            ok = (
                re.match(rf'^ISO{sep}\d', filename, re.IGNORECASE) or
                re.match(rf'^ISO{sep}IEC{sep}\S', filename, re.IGNORECASE) or
                re.match(rf'^ISO{sep}SAE(?:{sep}PAS)?{sep}\S', filename, re.IGNORECASE) or
                re.match(rf'^ISO{sep}(TR|TS){sep}\S', filename, re.IGNORECASE)
            )
            if not ok:
                return False

        # LN/MS/NAS/HST/MEP/NE/SPM/TAN — базовая проверка начала
        if folder_up in ('LN', 'MS', 'NAS', 'HST', 'MEP', 'NE', 'SPM', 'TAN'):
            if not re.match(rf'^(?i:{folder_up})\s*[\w-]', filename):
                return False

        # BS — разрешаем BS EN ISO, BS EN или BS <номер>
        if folder_up == 'BS':
            if not (re.match(r'^(?i:BS)\s*(EN\s*(ISO)?)', filename) or re.match(r'^(?i:BS)\s*\d', filename) or re.match(r'^(?i:BSEN)', filename)):
                return False

        return True
