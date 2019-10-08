# -*- coding: utf-8 -*-

from enum import Enum

class flags(Enum):
    START = 0      # 開始
    END = 1        # 終了
    ERROR = -1     # 問題発生
    CRASH = -2     # 処理破壊
    END_FORCED = -3# 強制終了