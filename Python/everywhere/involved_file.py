# -*- coding: utf-8 -*-

import os
from chardet.universaldetector import UniversalDetector

def createFolder(folderPass : str) -> bool:
    '''createFolder
    フォルダを新規作成する。
    
    Parameters
    ----------
    folderPass : str
        新規作成するフォルダのパス情報
    
    Returns
    ----------
    bool
        True(成功または、すでに存在している) , False(失敗)
    '''
    if os.path.isdir(folderPass): return True
    try:
        os.makedirs(folderPass)
        return True
    except FileExistsError as e:
        print(e.strerror)  # エラーメッセージ
        print(e.errno)     # エラー番号
        print(e.filename)  # 作成できなかったディレクトリ名
    #多分通らないと思う。
    return False

def checkEncoding(fileName : str) -> str:
    '''checkEncoding
    文字コードの自動判定
    
    Parameters
    ----------
    fileName : str
        ファイルのパス
    
    Returns
    ----------
    encoding : str
        文字コードの文字列
    '''
    detector = UniversalDetector()
    with open(fileName, mode='rb') as f:
        for binary in f:
            detector.feed(binary)
            if detector.done:
                break
    detector.close()
    return detector.result['encoding']

def fileNameCheck(fileName : str,extension : str) -> str:
    '''fileNameCheck
    セーブする際、しっかりとした拡張子形式になっているかの確認
    
    Parameters
    ----------
    fileName : str
        ファイル名
    extension : str
        拡張子 , (例 : txt)
        ドットは要らないよ。
    
    Returns
    ----------
    fileName : str
        ファイル名
    '''
    names = fileName.split('.')
    if names[len(names) - 1] == extension:
        return fileName
    else:
        return fileName + '.' + extension