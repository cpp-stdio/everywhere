# -*- coding: utf-8 -*-

import time
from enum import Enum

class STATUSES(Enum):
    START = 0       # 開始
    MAIN = 1        # 主要処理実行中
    END = 2         # 終了
    ERROR = -1      # 問題発生
    CRASH = -2      # 処理破壊
    END_FORCED = -3 # 強制終了

class PROCESS(Enum):
    EXECUTING = 0 # 実行中
    ENDED = 1     # 終了後
    ERROR = 2     # エラー

class needInThread:
    '''
    マルチスレッドが必要な場合に継承することで、マルチスレッド処理を簡易なものにする為のクラス。
    Pythonはマルチスレッドを実行するクラス、関数等が多く存在する為、このクラス内では用意はしない。
    一番簡単なマルチスレッド用のクラスは、
    import threading
    thread = threading.Thread(関数名)
    '''        
    def __init__(self):
        self._status = STATUSES.END
        self.__beginTime = time.time()
        self.__endTime = time.time()
    
    def _executeBegin(self):
        '''
        マルチスレッドの開始を宣言する為の処理
        '''
        self.__beginTime = time.time()
        self._status = STATUSES.START
        
    def _executeEnd(self, status : STATUSES):
        '''
        マルチスレッドの終了を宣言する為の処理
        '''
        self.__endTime = time.time()
        self._status = status
    
    def executedTime(self) -> float:
        '''
        マルチスレッド処理に掛かった時間。
        
        Returns
        ----------
        time : float
            秒単位
        '''
        return self.__endTime - self.__beginTime
    #--------------------------------------------------------------
    #   GET & SET & CHECK
    #--------------------------------------------------------------
    def getBeginTime(self):
        return self.__beginTime
    
    def getEndTime(self):
        return self.__endTime
    
    def procesCheck(self) -> STATUSES:
        # 現在実行中です。
        if self._status in [STATUSES.START, STATUSES.MAIN, STATUSES.CRASH]:
            return PROCESS.EXECUTING
        
        # 終了後です。
        if self._status in [STATUSES.END, STATUSES.END_FORCED]:
            return PROCESS.ENDED
        
        # 問題が発生しました。
        if self._status == STATUSES.ERROR:
            return PROCESS.ERROR
