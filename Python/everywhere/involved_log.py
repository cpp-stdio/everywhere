# -*- coding: utf-8 -*-

import time
import threading
import datetime
from enum import Enum

#自身のディレクトリを表示する
import sys,pathlib
__directoryName = str(pathlib.Path(__file__).resolve().parent)
sys.path.append(__directoryName)
import involved_file
import involved_other

class modeType(Enum):
    debug = 0     # デバック用
    release = 1   # リリース用
    tool = -0xFF  # ツール専用で、デバックとかリリースとかの概念が存在しないモード

class debugLog:
    
    def __init__(self,filePath : str, mode = modeType.debug, autoSave = False):
        '''
        コンストラクタ
        
        Parameters
        ----------
        filePath : str
            保存する、ファイルパス。
        mode : modeType
            モードによるログの制御
        autoSve : bool
            自動保存を実行するか、はい(True),いいえ(False)
        '''
        self.__texts = []
        self.__filePath = filePath
        self.__mode = mode
        self.__autoSave = autoSave
        #自動保存を実行する
        if self.__autoSave:
            self.__thread = threading.Thread(target = self.__autoSaveFunction)
            self.__thread
            self.__thread.start()
        
    def __delete__(self):
        self.release()
        
    def release(self):
        '''
        自動保存機能をONにした場合は必ず呼んで下さい。
        '''
        if self.__autoSave:
            self.__autoSave = False
            self.__thread.join()
        self.save()
        
    def add(self,message : str):
        '''
        ファイル、関数、行番号を取得した後、ログに書き込む
        
        Parameters
        ----------
        message : str
            プログラムからのメッセージ
        location : str
            ファイル、関数、行番号を取得関数を動かす為、引数に入れているだけで、
            常時、未入力で良い。
        '''
        if self.__mode == modeType.release: return
        self.__texts.append(message + '\n')
        
    def save(self, fileName = ''):
        '''
        ログを保存する、保存文字コードは'utf-8'
        
        Parameters
        ----------
        fileName : str
            保存するファイル名を変更したい場合、ファイル名を入力
            未入力の場合は、self.__filePathに登録している階層に'yyyy-mm-dd-hh-mm-ss.xxx.log'で保存される。
        '''
        if self.__mode == modeType.release: return
        
        if not fileName:
            if not self.__filePath: return
            fileName = self.__filePath +'\\' + str(datetime.datetime.now()).replace(':','-') + '.log'
        else:
            fileName = involved_file.fileNameCheck(fileName,'log')
        
        with open(fileName, mode='w',encoding = 'utf-8') as f:
            f.writelines(self.__texts)
            
    def __autoSaveFunction(self):
        '''
        自動保存
        '''
        while(self.conditions()):
            try:
                #30秒くらい待つ
                involved_other.sleep(30.0)
                #スリーブした後なのでもう一回条件を調べる
                if not self.conditions(): break
                #保存処理開始
                fileName = self.__filePath +'\\autoseve'
                if involved_file.createFolder(fileName):
                    fileName += '\\' + str(datetime.datetime.now()).replace(':','-') + '.log'
                    self.seve(fileName)
                else:
                    break
            except:
                break
        #確定でこの変数をFalseにする為
        self.__autoSave = False
        
    def conditions(self) -> bool:
        '''
        自動保存内にて下記の条件に当てはまっていれば即終了する
        '''
        try:
            if not self.__autoSve: return False
            if self.__mode == modeType.release: return False
            if not self.__filePath: return False
        except:
            return False
        return True
    
    def delete(self):
        '''
        メモリに大量にメッセージを残すのは危険な気もする為のログ情報の削除を行う
        '''
        self.text = []
        
    #--------------------------------------------------------------
    #   GET & SET
    #--------------------------------------------------------------
    def setMode(self,mode : modeType):
        self.__mode = mode