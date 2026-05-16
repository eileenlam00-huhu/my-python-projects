import os
import cv2
import numpy as np
from .screenshot import safe_imread, safe_cv_imread
from .utils import logger


def check_ui_overflow(base_img_path, test_img_path):
    try:
        if not base_img_path or not test_img_path:
            return 'N/A'
        if not os.path.exists(base_img_path) or not os.path.exists(test_img_path):
            return 'N/A'

        img_base = safe_imread(base_img_path, 0)
        img_test = safe_imread(test_img_path, 0)
        if img_base is None or img_test is None:
            return 'N/A'

        _, thresh_base = cv2.threshold(img_base, 50, 255, cv2.THRESH_BINARY)
        _, thresh_test = cv2.threshold(img_test, 50, 255, cv2.THRESH_BINARY)
        rows, cols = thresh_base.shape
        diff = cv2.absdiff(thresh_base, thresh_test)
        change_ratio = np.count_nonzero(diff) / (rows * cols)
        return 'Potential Overflow' if change_ratio > 0.05 else 'Normal'
    except Exception as e:
        logger(f'[COMPARE ERROR] {e}')
        return 'Check Error'


def analyze_ui_diff(base_path, test_path):
    if not base_path or not test_path:
        return 'Missing Path'
    img_b = safe_cv_imread(base_path, 0)
    img_t = safe_cv_imread(test_path, 0)
    if img_b is None or img_t is None:
        logger(f'[WARN] 图片读取失败 base={base_path}, test={test_path}')
        return 'Read Error'
    if img_b.shape != img_t.shape:
        logger('[WARN] 图片尺寸不一致，自动resize')
        img_t = cv2.resize(img_t, (img_b.shape[1], img_b.shape[0]))
    _, thresh_b = cv2.threshold(img_b, 70, 255, cv2.THRESH_BINARY)
    _, thresh_t = cv2.threshold(img_t, 70, 255, cv2.THRESH_BINARY)
    diff = cv2.absdiff(thresh_b, thresh_t)
    diff_count = np.count_nonzero(diff)
    total_pixels = diff.size
    if total_pixels == 0:
        return 'Bad Image'
    diff_ratio = (diff_count / total_pixels) * 100
    return f'Warning({diff_ratio:.1f}%)' if diff_ratio > 3.0 else 'Pass'
