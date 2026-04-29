import cv2
import numpy as np
from .constants import LANG_PAGE_NEXT_BTN
from .click import click_with_evtest
from .screenshot import safe_imread
from .utils import logger

ALL_LANGUAGES = {
    '1': ('中文（CN）', (140, 529)),
    '2': ('English', (128, 387)),
    '3': ('Deutsch', (131, 207)),
    '4': ('Español', (226, 573)),
    '5': ('Français', (218, 380)),
    '6': ('Italiano', (205, 209)),
    '7': ('Português', (313, 557)),
    '8': ('Русский', (313, 380)),
    '9': ('Turkish', (320, 217)),
    '10': ('日本語', (402, 556)),
    '11': ('한국어', (407, 364)),
    '12': ('العربية', (408, 213)),
    '13': ('繁体中文', (140, 529)),
    '14': ('Polski', (111, 394)),
    '15': ('Tiếng Việt', (131, 207)),
    '16': ('Bahasa Indonesia', (226, 573)),
    '17': ('ไทย', (218, 380)),
    '18': ('Bahasa Melayu', (205, 209)),
    '19': ('עברית', (313, 557)),
    '20': ('Afrikaans', (313, 380))
}


def select_languages():
    print('\n' + '=' * 40)
    print('      20国语言多选清单')
    print('=' * 40)
    for k, v in ALL_LANGUAGES.items():
        print(f"{k.rjust(2)}. {v[0]}")
    print('=' * 40)
    choice = input("请输入序号多选(如 1,2,5) 或输入 'all': ").strip().lower()

    selected_data = []
    if choice == 'all':
        for k, v in ALL_LANGUAGES.items():
            selected_data.append((int(k), v[0], v[1]))
    else:
        for idx in choice.split(','):
            idx = idx.strip()
            if idx in ALL_LANGUAGES:
                info = ALL_LANGUAGES[idx]
                selected_data.append((int(idx), info[0], info[1]))

    print(f'--- 已确认选择 {len(selected_data)} 种语言 ---')
    return selected_data


def check_language_switch(base_path, test_path):
    img_base = safe_imread(base_path, 0)
    img_test = safe_imread(test_path, 0)

    if img_base is None or img_test is None:
        return False, 0

    h, w = img_base.shape
    y1, y2 = int(h * 0.15), int(h * 0.90)
    x1, x2 = int(w * 0.20), int(w * 0.80)

    img_base = img_base[y1:y2, x1:x2]
    img_test = img_test[y1:y2, x1:x2]

    _, b1 = cv2.threshold(img_base, 80, 255, cv2.THRESH_BINARY)
    _, b2 = cv2.threshold(img_test, 80, 255, cv2.THRESH_BINARY)
    diff = cv2.absdiff(b1, b2)

    diff_ratio = np.count_nonzero(diff) / diff.size
    ok = diff_ratio > 0.005
    return ok, diff_ratio
