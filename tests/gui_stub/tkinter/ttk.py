from tkinter import _W,_ids
class Style(_W):
    def theme_use(self,*a,**k): pass
    def map(self,*a,**k): pass
class Frame(_W): pass
class Combobox(_W): pass
class Checkbutton(_W): pass
class Scrollbar(_W): pass
class Progressbar(_W): pass
class Treeview(_W):
    def __init__(self,*a,**k): super().__init__(); self._items=[]
    def heading(self,*a,**k): pass
    def column(self,*a,**k): pass
    def insert(self,parent,idx,values=(),tags=()):
        _ids[0]+=1; iid="I%d"%_ids[0]; self._items.append(iid); return iid
    def get_children(self): return list(self._items)
    def delete(self,iid):
        if iid in self._items: self._items.remove(iid)
    def set(self,*a,**k): pass
    def item(self,*a,**k): pass
    def detach(self,iid):
        if iid in self._items: self._items.remove(iid)
    def reattach(self,iid,p,idx):
        if iid not in self._items: self._items.append(iid)
    def see(self,iid): pass
    def tag_configure(self,*a,**k): pass
    def identify_region(self,*a): return "cell"
    def identify_column(self,*a): return "#1"
    def identify_row(self,*a): return self._items[0] if self._items else ""
