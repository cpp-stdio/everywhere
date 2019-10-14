# -*- coding: utf-8 -*-
"""いつもつかう(everywhere)
    
    Created on Thu Jan 24 10:10:43 2019
    Updated on wed Oct 02 15:51:00 2019
     
    * 毎回よく使うので、このような形で保存する。
    * We will use. it is everywhere.
    
Todo:
    Spyderというエディターを使用しています。
    We are using Spyder. it is editor.
    
    未使用って警告が出ていますが別に問題ありません。
    warning "imported but unused". But no problem.
    
Examples:
    import sys, pathlib
    __directoryName = str(pathlib.Path(__file__).resolve().parent)
    sys.path.append(__directoryName)
    import everywhere
    
    everywhere.file.createFolder('C:\\Users\\Public\\test')
"""
from everywhere import involved_file as file
from everywhere import involved_other as other
from everywhere import involved_search as search
from everywhere import involved_thread as thread
from everywhere import involved_log as log