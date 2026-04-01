#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import subprocess
import time
from datetime import datetime
import random


class LightEffectStressTest:
    def __init__(self):
        self.config = {
            "touch_device": "/dev/input/event0",

            "effect_coords": {
                "effect_1": [164, 556],
                "effect_2": [170, 436],
                "effect_3": [172, 325],
                "effect_4": [160, 212],
                "effect_5": [168, 116],
                "effect_6": [263, 566],
                "effect_7": [252, 437],
            },

            # ✅ 进入灯效页面（只执行一次）
            "enter_steps": [
                ("进入设置项", [370, 259]),
                ("点击打印设置", [40, 326]),
                ("点击状态灯选项", [149, 194]),
                ("点击灯效样式", [242, 260]),
            ],

            "enter_interval": 0.4,#进入页面时点击间隔
            "switch_interval": 1.0,#切换灯效的速度
            "round_interval": 0.8,#每轮之间的停顿

            "light_effects": [
                "柔和白光",
                "彩虹灯效横向流水",
                "彩虹灯效竖向流水",
                "粉蓝灯效",
                "青色灯效",
                "蓝紫灯效",
                "粉色灯效",
            ]
        }

        self.stats = {
            "total": 0,
            "success": 0,
            "fail": 0,
            "rounds": 0,
            "start_time": None
        }

    def log(self, msg):
        now = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"[{now}] {msg}")

    # ✅ 已改为 touch_control
    def click(self, x, y):
        device = self.config["touch_device"]

        # ✅ 随机偏移（±3像素）
        x_offset = random.randint(-3, 3)
        y_offset = random.randint(-3, 3)

        real_x = x + x_offset
        real_y = y + y_offset

        cmd = f"/touch_control {device} {real_x} {real_y}"

        self.log(f"👉 点击: ({real_x}, {real_y})")

        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)

        if result.returncode != 0:
            self.log(f"❌ 点击失败: {result.stderr.strip()}")
            return False

        # ✅ 更短点击间隔 + 随机抖动
        time.sleep(0.05 + random.uniform(0, 0.05))
        return True

    # ✅ 只执行一次
    def enter_page_once(self):
        self.log("🚀 进入灯效页面（仅一次）")

        for i, (name, coord) in enumerate(self.config["enter_steps"]):
            x, y = coord
            self.log(f"步骤 {i + 1}: {name}")

            if not self.click(x, y):
                self.log(f"❌ {name} 失败")
                return False

            time.sleep(self.config["enter_interval"])

        time.sleep(1.0)
        self.log("✅ 已进入灯效页面")
        return True

    def switch_effect(self, index):
        effect_name = self.config["light_effects"][index]
        coord_key = f"effect_{index + 1}"

        x, y = self.config["effect_coords"][coord_key]

        self.log(f"[{index + 1}] {effect_name}")

        if not self.click(x, y):
            return False

        time.sleep(self.config["switch_interval"])
        return True

    # ✅ 主循环（7个灯效=1轮）
    def run_loop(self, rounds=None, minutes=None):
        if not rounds and not minutes:
            self.log("⚠️ 请指定 -r 或 -d")
            return

        self.stats["start_time"] = time.time()
        end_time = time.time() + minutes * 60 if minutes else None

        try:
            while True:
                self.stats["rounds"] += 1
                self.log(f"🔁 第 {self.stats['rounds']} 轮")

                # ✅ 随机顺序
                indices = list(range(len(self.config["light_effects"])))
                random.shuffle(indices)

                for i in indices:
                    self.stats["total"] += 1

                    if self.switch_effect(i):
                        self.stats["success"] += 1
                        self.log("✓ 成功")
                    else:
                        self.stats["fail"] += 1
                        self.log("✗ 失败")

                self.log(f"📊 总:{self.stats['total']} 成功:{self.stats['success']} 失败:{self.stats['fail']}")

                if rounds and self.stats["rounds"] >= rounds:
                    break

                if end_time and time.time() >= end_time:
                    break

                time.sleep(self.config["round_interval"])

        except KeyboardInterrupt:
            self.log("⛔ 用户中断")

        self.print_results()

    def print_results(self):
        elapsed = time.time() - self.stats["start_time"]

        print("\n======================")
        print("测试结果")
        print("======================")
        print(f"时长: {elapsed:.1f}s")
        print(f"轮数: {self.stats['rounds']}")
        print(f"总次数: {self.stats['total']}")
        print(f"成功: {self.stats['success']}")
        print(f"失败: {self.stats['fail']}")

        if self.stats["total"] > 0:
            rate = self.stats["success"] / self.stats["total"] * 100
            print(f"成功率: {rate:.2f}%")

        print("======================")


def main():
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument('-r', '--rounds', type=int)
    parser.add_argument('-d', '--duration', type=int)

    args = parser.parse_args()

    tester = LightEffectStressTest()

    # ✅ 只进一次页面
    if not tester.enter_page_once():
        print("❌ 进入页面失败")
        return

    # ✅ 开始循环
    tester.run_loop(rounds=args.rounds, minutes=args.duration)


if __name__ == "__main__":
    main()