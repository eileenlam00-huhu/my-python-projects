from .constants import HOME_RESET_STEPS
from .click import click_with_evtest
from .utils import logger

HOME_RESET_STEPS = HOME_RESET_STEPS

UI_STRUCTURE = {
    '首页': {
        'entry': (None, None),
        'depth': 0,
        'sub_pages': {
            '主界面': [('click', None, None)]
        }
    },
    '文件列表': {
        'entry': (191, 390),
        'depth': 1,
        'sub_pages': {
            '本地浏览': [None],
            '本地长按菜单': [('pulse', 176, 516)],
            '本地复制操作': [('click', 29, 449)],
            'U盘': [(49, 287)],
            '打印历史': [(25, 135)]
        }
    },
    '耗材页面': {
        'entry': (398, 546),
        'depth': 1,
        'sub_pages': {
            'CFS lite': [None],
            '料架': [(45, 291)]
        }
    },
    '控制页面': {
        'entry': (383, 414),
        'depth': 1,
        'sub_pages': {
            '控制': [None],
            'XYZ': [(69, 351)]
        }
    },
    '设置页面': {
        'entry': (406, 247),
        'depth': 1,
        'sub_pages': {
            '账号': [(None, None)],
            '打印设置': {
                'entry': (66, 315),
                'depth': 1,
                'sub_pages': {
                    '状态灯页面': {
                        'entry': (138, 354),
                        'depth': 2,
                        'sub_pages': {
                            '灯效样式': {
                                'entry': (228, 178),
                                'depth': 3,
                                'sub_pages': {}
                            }
                        }
                    }
                }
            },
            '系统': [(54, 108)]
        }
    },
    '帮助页面': {
        'entry': (408, 79),
        'depth': 1,
        'sub_pages': {
            '提示': [None],
            'WiKi': [(33, 334)],
            '维护': [(58, 160)]
        }
    }
}


def reset_to_home(page_name):
    depth = UI_STRUCTURE.get(page_name, {}).get('depth', 1)
    if depth <= 0:
        logger(f'  [SKIP] 页面 {page_name} 深度为 0，无需执行返回动作')
        return
    for i in range(depth):
        click_with_evtest(63, 617, page_label=f'返回首页_第{i + 1}步', retries=3)
        from .constants import SCREEN_RENDER
        import time
        time.sleep(2.0)
