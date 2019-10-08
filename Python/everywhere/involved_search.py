# -*- coding: utf-8 -*-

def searchList(value, layers : list, abate = 1) -> list:
    '''searchList
    レイヤー設計になっているlist(tuple)から情報を検索する為の関数
    
    Parameters
    ----------
    value : ???
        検索対象値
    layers : list(tuple)
        検索対象のlist(tuple)
    abate : int
        初期値 : 1個下まで検索
        正の数 : 数値に合わせた回数分、検索する
        負の数 : どの数字でも無限とみなされる
    
    Returns
    ----------
    returnArray : list
        検索された結果を一次元配列で返却する
    '''
    returnArray = []
    for layer in layers:
        if __equals(value,layer):
            returnArray.append(layer)
        else:
            array = __search(value, layer, abate)
            if array != None: returnArray.extend(array)
    return returnArray

def searchDict(value, layers : dict, abate = 1) -> list:
    '''searchDict
    レイヤー設計になっているdictから情報を検索する為の関数
    
    Parameters
    ----------
    value : ???
        検索対象値
    layers : list
        検索対象のlist
    abate : int
        初期値 : 1個下まで検索
        正の数 : 数値に合わせた回数分、検索する
        負の数 : どの数字でも無限とみなされる
    
    Returns
    ----------
    returnArray : list
        検索された結果を一次元配列で返却する
    '''
    returnArray = []
    for key, item in layers.items():
        if __equals(value,key):
            returnArray.append(item)
        elif __equals(value,item):
            returnArray.append(key)
        else:
            array = __search(value, item, abate)
            if array != None: returnArray.extend(array)
    return returnArray
#==============================================================
#
#   private
#
#==============================================================
def __equals(thisValue,thatValue) -> bool:
    #型と値をしっかりと検出する為、クラスの場合は対応できない。
    if type(thisValue) == type(thatValue) and thisValue == thatValue:
        return True
    return False

def __search(value, layer, abate) -> list:
    if abate == 0: return None # 0の場合のみ処理を終了させる。
    if abate < -1: abate = -1  # オーバーフローを回避する。
    # 型を判定した後、相互再起をしている。
    if type(layer) == list or type(layer) == tuple:
        array = searchList(value, layer, abate - 1)
        if len(array) >= 1: return array
    elif type(layer) == dict:
        array = searchDict(value, layer, abate - 1)
        if len(array) >= 1: return array
    return None
