# -*- coding: utf-8 -*-
"""
테스트 전용 가짜 tkinter (화면 없이 GUI 로직만 돌리기 위함). 실제 프로그램은 쓰지 않는다.
GUI에서 새 tkinter 기능을 쓰면 여기에도 흉내를 추가해야 할 수 있다.
"""
_ids=[0]
class _W:
    def __init__(self,*a,**k): self._d={}
    def __getattr__(self,n): return lambda *a,**k: None
    def __setitem__(self,k,v): self._d[k]=v
    def __getitem__(self,k): return self._d.get(k)
    def config(self,*a,**k): pass
    def configure(self,*a,**k): pass
    def pack(self,*a,**k): pass
    def grid(self,*a,**k): pass
    def bind(self,*a,**k): pass
    def geometry(self,*a,**k): return "1040x760+0+0"
    def title(self,*a,**k): pass
    def minsize(self,*a,**k): pass
    def protocol(self,*a,**k): pass
    def after(self,ms,fn=None,*a):
        if callable(fn): fn(*a)
    def update_idletasks(self): pass
class Tk(_W): pass
class Frame(_W): pass
class Label(_W): pass
class Entry(_W): pass
class Button(_W): pass
class Menu(_W):
    def add_command(self,*a,**k): pass
    def add_separator(self,*a,**k): pass
    def entryconfig(self,*a,**k): pass
class Toplevel(_W): pass
class StringVar(_W):
    def __init__(self,*a,value=None,**k): self._v=value or ""; self._cb=[]
    def get(self): return self._v
    def set(self,v):
        self._v=v
        for c in list(self._cb): c()
    def trace_add(self,mode,cb): self._cb.append(lambda *a: cb())
class BooleanVar(StringVar):
    def __init__(self,*a,value=False,**k): self._v=value; self._cb=[]
    def get(self): return self._v
TclError=Exception
