# -*- coding: utf-8 -*-
"""ファイル関連(involved_file)
    
    Created on Thu Jan 24 10:10:43 2019
    Updated on Tue Oct 08 15:56:00 2019
     
    * ログを出力します。不具合の解析又は単体テストにも使用できます。
    * Output log and can use for debug or unit test.
    
Examples:
    logs = everywhere.log.log('C:\\Users\\Public\\test', mode = everywhere.log.modeType.debug, autoSave = True)
    
    # autoSaveの値がTrueの場合、必ずデストラクタを呼んで下さい。
    # if autoSave of True so you must to call destructor.
    del logs
"""
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
    debug = 0
    release = 1
    tool = -0xFF  # ツール専用で、デバックとかリリースとかの概念が存在しないモード

class log:
    
    def __init__(self,filePath : str, mode = modeType.debug, autoSave = False):
        '''
        コンストラクタ
        
        Parameters
        ----------
        filePath : str
            保存するファイルパス
            It's file path to save.
            
        mode : modeType
            モードによるログの制御
            operation control to logs
            
        autoSve : bool
            自動保存を実行するか、はい(True),いいえ(False)
            Do you save automatically? Yes(True),No(False)
        '''
        self.__texts = []
        self.__filePath = filePath
        self.__mode = mode
        self.__autoSave = autoSave
        #自動保存を実行する
        if self.__conditions():
            self.__thread = threading.Thread(target = self.__autoSaveFunction)
            self.__thread
            self.__thread.start()
        else:
            self.__autoSave = False
        
    def __delete__(self):
        self.release()
        
    def release(self):
        '''
        自動保存機能をONにした場合は必ず呼んで下さい。
        if autoSave of True so you must to call
        '''
        if self.__autoSave:
            self.__autoSave = False
            self.__thread.join()
        self.save()
        
    def add(self,message : str):
        '''
        ログに書き込む
        It's add to log
        
        Parameters
        ----------
        message : str
            プログラムからのメッセージ
            It's message from program
        '''
        if self.__mode == modeType.release: return
        self.__texts.append(message + '\n')
        
    def save(self, fileName = ''):
        '''
        ログを保存する、保存文字コードは'utf-8'
        It's save logs, character code is 'utf-8'
        
        Parameters
        ----------
        fileName : str
            保存するファイル名を変更したい場合、ファイル名を入力
            If you want to change the file name to save.
            未入力の場合は、self.__filePathに登録している階層に'yyyy-mm-dd-hh-mm-ss.xxx.log'で保存される。
            If not entering, 'self.__filePath' registration hierarchy for name 'yyyy-mm-dd-hh-mm-ss.xxx.log' to save.
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
        It's auto save
        '''
        while(self.__conditions()):
            try:
                #30秒くらい待つ
                involved_other.sleep(30.0)
                #スリーブした後なのでもう一回条件を調べる
                if not self.__conditions(): break
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
        
    def __conditions(self) -> bool:
        '''
        自動保存内にて下記の条件に当てはまっていれば即終了する
        If the following conditions are met in auto save, it will end at once.
        '''
        try:
            if not self.__autoSve: return False
            if self.__mode == modeType.tool: return False
            if self.__mode == modeType.release: return False
            if not self.__filePath: return False
        except:
            return False
        return True
    
    def delete(self):
        '''
        メモリに大量にメッセージを残すのは危険な気もする為のログ情報の削除を行う
        delete logs, that's dangerous to use much memory.
        '''
        self.__texts = []
        
    #--------------------------------------------------------------
    #   GET & SET
    #--------------------------------------------------------------
    def setMode(self,mode : modeType):
        self.__mode = mode