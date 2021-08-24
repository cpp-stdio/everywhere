# -*- coding: utf-8 -*-
import time

def betweenSplit(textStr : str,specificA : str,specificB : str) -> list:
    '''betweenSplit
    文字列の中にある、特定の文字列から特定の文字列までを取得する
    
    Parameters
    ----------
    textStr : str
        とある文字列
    specificA : str
        地点Aの文字列
    specificB : str
        地点Bの文字列
    
    Returns
    -------
    returnStrings : list[str]
        地点Aと地点Bの間の文字列、無かったら空の配列
    '''
    returnStrings = []
    #同じ文字列なら用途が違う
    if(specificA == specificB): return returnStrings
    textArray1 = textStr.split(specificA)
    
    for s in range(1,len(textArray1)):
        textArray2 = textArray1[s].split(specificB,1)
        if len(textArray2) >= 1:
            returnStrings.append(textArray2[0])
    
    return returnStrings

def is_value(textStr : str):
    '''is_value
    文字列を数字に変換する
    
    Parameters
    ----------
    textStr : str
        とある文字列
    
    Returns
    ----------
    OKflag : bool
        文字列を数字に変換、成功(True)、失敗(False)
    value : int,float,complex,str
        上記の型がどれになるのかは分からない。
    '''
    try:
        i = int(textStr)
        return True,i
    except:
        try:
            f = float(textStr)            
            return True,f
        except:
            try:
                c = complex(textStr)
                return True,c
            except:
                return False,textStr
    return False,textStr#要らな気がするが一応...

def sleep(second : float):
    '''wait
    Python既存のsleep関数をマルチスレッドで呼んだ場合、
    なぜか、シングルスレッドの方にも影響が出てしまう。
    原因が分からないので、一定時間プログラムを停止させるコードを手書きする。
    
    Parameters
    ----------
    second : float
        停止時間（秒）
    '''
    start = time.time()
    while(True): 
        if time.time() - start <= second: return
    
