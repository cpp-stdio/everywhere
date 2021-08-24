# -*- coding: utf-8 -*-# -*- coding: utf-8 -*-
"""ファイル関連(involved_file)
    
    Created on Thu Jan 24 10:10:43 2019
    Updated on Tue Oct 08 15:56:00 2019
     
    * ファイル関連の関数を纒ています。
    * It have file related only functions.
"""

import os
from chardet.universaldetector import UniversalDetector

def createFolder(folderPass : str) -> bool:
    '''createFolder
    フォルダを新規作成する。
    It's create folder.
    
    Parameters
    ----------
    folderPass : str
        新規作成するフォルダのパス情報
        folder pass information to create new.
    Returns
    ----------
    bool
        True(成功または、すでに存在している) , False(失敗)
        True(succeed or already exists) , False(failed).
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
    It's automatic determination of character code.
    
    Parameters
    ----------
    fileName : str
        ファイル名
        It's file name.
    
    Returns
    ----------
    encoding : str
        文字コードの文字列
        character encoding.
        
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
    When saving, It's checking that the extension format.
    
    Parameters
    ----------
    fileName : str
        ファイル名
        It's file name.
    extension : str
        拡張子 , (例 : txt)
        extension , (case : txt)
        ドットは要らないよ。
        Please don't need dots.
    
    Returns
    ----------
    fileName : str
        ファイル名
        It's file name.
    '''
    names = fileName.split('.')
    if names[len(names) - 1] == extension:
        return fileName
    else:
        return fileName + '.' + extension