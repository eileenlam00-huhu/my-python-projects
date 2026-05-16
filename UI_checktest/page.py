from .constants import HOME_RESET_STEPS
from .click import click_with_evtest
from .utils import logger

HOME_RESET_STEPS = HOME_RESET_STEPS


def _make_action(action_type, x=None, y=None):
    return {
        'type': action_type,
        'x': x,
        'y': y,
    }


def click_action(x=None, y=None):
    return _make_action('click', x, y)


def pulse_action(x, y):
    return _make_action('pulse', x, y)


def noop_action():
    return [_make_action('noop', None, None)]


def build_page(entry=None, depth=1, sub_pages=None):
    return {
        'entry': entry,
        'depth': depth,
        'sub_pages': sub_pages or {}
    }


def is_page_definition(value):
    return isinstance(value, dict) and 'entry' in value and 'sub_pages' in value


def normalize_actions(actions):
    if actions is None:
        return noop_action()
    if isinstance(actions, dict) and not is_page_definition(actions):
        return [actions]
    if isinstance(actions, tuple):
        if len(actions) == 3:
            _, x, y = actions
        else:
            x, y = actions
        return [click_action(x, y)]
    if isinstance(actions, list):
        normalized = []
        for item in actions:
            if item is None:
                normalized.extend(noop_action())
            elif isinstance(item, dict):
                normalized.append(item)
            elif isinstance(item, tuple):
                if len(item) == 3:
                    _, x, y = item
                else:
                    x, y = item
                normalized.append(click_action(x, y))
            else:
                normalized.extend(noop_action())
        return normalized
    return [actions]


def get_action_coords(action):
    if not isinstance(action, dict):
        return None, None
    return action.get('x'), action.get('y')


def get_action_type(action):
    if not isinstance(action, dict):
        return 'noop'
    return action.get('type', 'noop')


UI_STRUCTURE = {
    '首页': build_page(entry=(None, None), depth=0, sub_pages={
        '主界面': noop_action()
    }),
    '文件列表': build_page(entry=(191, 390), depth=1, sub_pages={
        '本地浏览': noop_action(),
        '本地长按菜单': [pulse_action(176, 516)],
        '本地复制操作': [click_action(29, 449)],
        'U盘': [click_action(49, 287)],
        '打印历史': [click_action(25, 135)]
    }),
    '耗材页面': build_page(entry=(398, 546), depth=1, sub_pages={
        'CFS lite': noop_action(),
        '料架': [click_action(45, 291)]
    }),
    '控制页面': build_page(entry=(383, 414), depth=1, sub_pages={
        '控制': noop_action(),
        'XYZ': [click_action(69, 351)]
    }),
    '设置页面': build_page(entry=(406, 247), depth=1, sub_pages={
        '账号': noop_action(),
        '打印设置': build_page(entry=(66, 315), depth=1, sub_pages={
            '状态灯页面': build_page(entry=(138, 354), depth=2, sub_pages={
                '灯效样式': noop_action()
            })
        }),
        '系统': [click_action(54, 108)]
    }),
    '帮助页面': build_page(entry=(408, 79), depth=1, sub_pages={
        '提示': noop_action(),
        'WiKi': [click_action(33, 334)],
        '维护': [click_action(58, 160)]
    })
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
