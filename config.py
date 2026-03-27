import re

regular_num = re.compile("\\d+[.,]?\\d+")                                 # Номер

number_section = r'\d+(?:[,.]\d+)?'
regular_cableSection = re.compile(rf"\d?[(]?{number_section}[xх×]{number_section}[)]?(?:[xх×]{number_section})*")  # Сечение
regular_cableMark = re.compile("[н][г][' ']?[(/]?[AАС]?[)/]?[--]")                  # Марка кабеля 2

arrayCableMarks = ['КУГПЭПнг', 
                   'SM-MLT', 
                   'КППГЭнг', 
                   'эп(А)-', 
                   'нг(А)', 
                   'КГН', 
                   'КРА6.560.106-27', 
                   'EC4-50-HF', 
                   'TППэп(A)-HДГ',
                   'СПЕЦЛАН',
                   'Complete',
                   'UNITRONIC',
                   'ТАДУ',
                   'KCнг(A)-HF',
                   '6XV1871-2F',
                   'Hyperline',
                   'СКАБ',
                   'PK75',
                   'PROFINET',
                   'F/UTP',
                   'FO-ST-IN/OUT',
                   'КАГЭнг(B)',
                   'нг(A)-HF',
                   'LLMD',
                   'ТК-1, тип Б исп.45',
                   'Спецкабель',
                   'ТППэп',
                   'нг(A)',
                   'КПЭТИнг(B)-HF-Т',
                   'РКГМ',
                   'ППГнг(A)-HF',
                   'ОКБнг(В)-',
                   'OKБнг(B)-',
                   'ОБР-У-4ОВ SM',
                   'КУГПвЭПвнг(A)-HF-Т',
                   'КПЭТИнг(B)-FRHF-Т',
                   'RG58',
                   'RG213'
                   ]

regular_KKS = re.compile("\\d{2,3}[А-ЯA-Z]{2,3}\\d{2,3}")                    # ККС оборудования 10BAT01GH301

# regular_KKS_any = re.compile(r'[A-Z0-9]{2,30}.?[A-Z0-9]{2,30}')  # ККС оборудования или помещения
regular_KKS_any = re.compile(r'(?=.*[A-Z])(?=.*\d)[A-Z\d\s.-]{5,}')   # ККС оборудования или помещения

# test
# часть строки
regular_KKS_room = re.compile(r'\d{2}[A-Z]{3}\d{1,3}R\d{1,3}')  # ККС помещения
regular_KKS_building = re.compile(r'\d{2}[A-Z]{3}')  # ККС здания
regular_KKS_equipment = re.compile(r'(?=.*[A-Z])(?=.*\d)(?!.*-.*-)(?!.*\s.*\s)[A-Z\d\s.-]{5,}') # ККС оборудования
# вся строка
regular_KKS_room_full = re.compile(r'^\d{2}[A-Z]{3}\d{1,3}R\d{1,3}$')  # ККС помещения
regular_KKS_building_full = re.compile(r'^\d{2}[A-Z]{3}$')  # ККС здания
regular_KKS_equipment_full = re.compile(r'^(?=.*[A-Z])(?=.*\d)(?!.*-.*-)(?!.*\s.*\s)[A-Z\d\s.-]{5,}$') # ККС оборудования

regular_axis_full = re.compile(r'^[+-]?\d{1,8}\.?,?\d{0,4}\.?$')   # координата целиком
regular_axis = re.compile(r'[+-]?\d{1,8}\.?,?\d{0,4}\.?')   # координата как часть текста


# --------- регулярное выражение журнала ---------
# Общая часть
common_start = r'([А-ЯA-Z]{3})\.([0-9]{4})\.([0-9]{2})'
# Вариативная часть
type1 = r'([А-ЯA-Z]{3})\.([А-ЯA-Z]{3})\.([А-ЯA-Z]{2})\.([А-ЯA-Z]{2})'        # UFC.CYY.SS.MB
type2 = r'([А-ЯA-Z]{3})\.([0-9]{1})\.([А-ЯA-Z]{2})\.([А-ЯA-Z]{2})'            # UKS.0.AP.MB
type3 = r'([0-9]{1})\.([А-ЯA-Z]{3})\.([А-ЯA-Z]{2})\.([А-ЯA-Z]{2})'            # 0.UKS.AP.MB
# Конец
number_part = r'([0-9]{4})'
suffix = rf'[-_]([А-ЯA-Z]{{3}})([0-9]{{4}})'  # необязательный суффикс
# Собираем всё вместе
regular_Journal = re.compile( rf'{common_start}(?:{type1}|{type2}|{type3}){number_part}(?:{suffix})?' )
# --------- регулярное выражение журнала ---------


regular_letter_minus = re.compile(r'[a-zA-Zа-яА-ЯёЁ-]')  # хотя бы одна буква или знак -