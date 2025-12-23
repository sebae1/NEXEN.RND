from __future__ import annotations

import wx
import wx.dataview as DV
from typing import Hashable

class TreeListNode:
    def __init__(self, parent: TreeListNode, key: Hashable, item: any):
        """
        Args:
            parent
                논리적 루트 노드 외에는 반드시 부모 노드를 가져야함.
            key
                ID 외에 이 노드를 가르킬 hashable key.
            item
                노드를 표현할 아이템.
        """
        self.parent = parent
        self.key = key
        self.item = item
        self.children: list[TreeListNode,] = []

    def get_level(self) -> int:
        level = 0
        cur = self.parent
        while cur:
            level += 1
            cur = cur.parent
        return level

class TreeListModelBase(DV.PyDataViewModel):
    def __init__(self, n_columns: int):
        """DV.DataView에 맞게 이 클래스를 상속하여 새로 정의되어야 함"""
        DV.PyDataViewModel.__init__(self)
        self.n_columns = n_columns
        self.nodes: dict[int, TreeListNode] = {} # {node_id (int): TreeListNode,}
        self.key_vs_id: dict[str, int] = {} # {key (Hashable): node_id (int),}
        self.logical_root = TreeListNode(None, None, None) # 표현되지 않는 논리적 루트 노드
        self.nodes[id(self.logical_root)] = self.logical_root

    def get_view_item(self, node: TreeListNode):
        item = DV.DataViewItem() if node is self.logical_root else DV.DataViewItem(id(node))
        return item

    def purge_subtree(self, node: TreeListNode):
        """dict/key 맵에서 node와 모든 후손 제거"""
        stack = [node]
        while stack:
            cur = stack.pop()
            # 자식들 먼저 스택에 넣고
            stack.extend(cur.children)
            # 맵 제거
            nid = id(cur)
            self.nodes.pop(nid, None)
            if cur.key is not None:
                self.key_vs_id.pop(cur.key, None)
            # 참조 끊기
            cur.children.clear()
            cur.parent = None

    def GetColumnCount(self):
        """컬럼 수 반환"""
        return self.n_columns
        
    def GetColumnType(self, col):
        """모든 컬럼을 string 타입으로 반환"""
        return "string"
        
    def GetValue(self, item, col):
        """노드로부터 값 표출을 어떻게 할 것인지 정의 필요"""
        raise NotImplementedError

    def GetChildren(self, parent, children):
        if not parent:
            for child_node in self.logical_root.children:
                children.append(DV.DataViewItem(id(child_node)))
        else:
            node_id = int(parent.GetID())
            if node_id in self.nodes:
                node = self.nodes[node_id]
                for child_node in node.children:
                    children.append(DV.DataViewItem(id(child_node)))
        return len(children)
        
    def IsContainer(self, item):
        if not item:
            return True
        node_id = int(item.GetID())
        if node_id in self.nodes:
            return len(self.nodes[node_id].children) > 0
        return False

    def HasContainerColumns(self, item):
        """컨테이너 노드도 모든 컬럼에 값을 가질 수 있음을 명시"""
        return True
        
    def GetParent(self, item):
        if not item:
            return DV.DataViewItem()
        
        node_id = int(item.GetID())
        if node_id in self.nodes:
            parent_node = self.nodes[node_id].parent
            if parent_node and parent_node != self.logical_root:
                return DV.DataViewItem(id(parent_node))
        
        return DV.DataViewItem()

    def GetAttr(self, item, col, attr):
        """셀의 속성 (색상, 폰트 등) 설정"""
        return False

class TreeListCtrl(DV.DataViewCtrl):
    def __init__(self, parent: wx.Window, model: TreeListModelBase, columns: dict[str, int]):
        """
        Args:
            columns: {label (str): width (int),}
        """
        DV.DataViewCtrl.__init__(
            self,
            parent,
            size=(sum(columns.values())+25, -1),
            style=DV.DV_ROW_LINES|DV.DV_HORIZ_RULES
        )
        self.model = model
        self.AssociateModel(self.model)
        for i, (label, width) in enumerate(columns.items()):
            renderer = DV.DataViewTextRenderer()
            col = DV.DataViewColumn(label, renderer, i, width=width)
            col.SetSortable(True)
            self.AppendColumn(col)

    def clear_nodes(self):
        """Logical root를 제외한 모든 최상위 노드 제거"""
        logical_root = self.model.logical_root
        self.Freeze()
        try:
            self.UnselectAll()
            top_nodes = list(logical_root.children)
            if not top_nodes:
                return
            logical_root.children.clear()
            parent_item = DV.DataViewItem()
            for n in top_nodes:
                item = DV.DataViewItem(id(n))
                self.model.ItemDeleted(parent_item, item)
                self.model.purge_subtree(n)
            self.model.nodes.clear()
            self.model.key_vs_id.clear()
            self.model.nodes[id(logical_root)] = logical_root
        finally:
            self.Thaw()

    def get_node_by_key(self, key: Hashable) -> TreeListNode:
        return self.model.nodes[self.model.key_vs_id[key]]

    def move_node(self, node: TreeListNode, down: bool):
        """node를 같은 부모 안에서 한 칸 위/아래로 이동"""
        # 0) 부모/형제 범위 파악
        parent = node.parent
        if parent is None:
            return
        siblings = parent.children
        try:
            i = siblings.index(node)
        except ValueError:
            return

        j = i + (1 if down else -1)
        if j < 0 or j >= len(siblings):
            return # 끝이라 못 움직임

        # 1) 순서 교환(모델 데이터 먼저)
        siblings[i], siblings[j] = siblings[j], siblings[i]

        # 2) 정렬(헤더) 켜져 있으면 수동 재배치가 보이지 않을 수 있음
        #    → 수동 위치를 쓰고 싶으면 헤더 정렬을 끄거나, '순서' 키로 Compare 구현을 쓰세요.
        #    여기서는 단순히 reseat로 강제 갱신
        parent_item = self.model.get_view_item(parent)
        item = DV.DataViewItem(id(node))

        if parent_item.IsOk():
            self.Expand(parent_item)

        self.Freeze()
        try:
            self.model.ItemDeleted(parent_item, item)
            self.model.ItemAdded(parent_item, item)
            self.EnsureVisible(item)
            self.UnselectAll()
            self.Select(item)
            self.SetFocus()
            self.Expand(item)
        finally:
            self.Thaw()

    def add_node(self, parent_node: TreeListNode|None, key: Hashable, item: any) -> TreeListNode:
        """노드 추가.
        parent_node=None으로 명시하면 논리적 루트 노드를 부모 노드로 함.
        표현상 최상위 노드가 됨.
        """
        if parent_node is None:
            parent_node = self.model.logical_root
        node = TreeListNode(parent_node, key, item)
        node_id = id(node)
        self.model.nodes[node_id] = node
        self.model.key_vs_id[key] = node_id
        parent_node.children.append(node)
        parent_item = self.model.get_view_item(parent_node)
        self.model.ItemAdded(
            parent_item,
            DV.DataViewItem(node_id)
        )
        if len(parent_node.children) == 1:
            self.Expand(parent_item)
        return node

    def delete_node(self, node: TreeListNode):
        if node is self.model.logical_root:
            raise ValueError
        try:
            node.parent.children.remove(node)
        except ValueError:
            pass
        self.model.ItemDeleted(
            self.model.get_view_item(node.parent),
            DV.DataViewItem(id(node))
        )
        self.model.purge_subtree(node)
    
    def update_node(self, node: TreeListNode):
        self.model.ItemChanged(DV.DataViewItem(id(node)))

    def expand_node(self, node: TreeListNode, flag: bool):
        item = DV.DataViewItem(id(node))
        if flag:
            self.Expand(item)
        else:
            self.Collapse(item)

    def reveal_and_select(self, node: TreeListNode):
        """노드를 화면에 보이게 하고 선택 상태로 만듦"""
        parent = node.parent
        logical_root = self.get_logical_root()
        while parent != logical_root:
            self.Expand(DV.DataViewItem(id(parent)))
            parent = parent.parent
        item = DV.DataViewItem(id(node))
        self.EnsureVisible(item)
        self.UnselectAll()
        self.Select(item)
        self.SetFocus()

    def get_logical_root(self) -> TreeListNode:
        return self.model.logical_root

    def get_selected_node(self) -> TreeListNode|None:
        item = self.GetSelection()
        if not item.IsOk():
            return
        node = self.model.nodes[int(item.GetID())]
        return node

    def get_selected_nodes(self) -> list[TreeListNode,]:
        nodes = []
        items = self.GetSelections()
        for item in items:
            if not item.IsOk():
                continue
            node_id = int(item.GetID())
            node = self.model.nodes[node_id]
            nodes.append(node)
        return nodes

    def OnItemActivated(self, event):
        """항목 더블클릭 이벤트"""
        item = event.GetItem()
        if item:
            name = self.model.GetValue(item, 0)
            wx.MessageBox(f"Activated: {name}", "Item Activated", wx.OK | wx.ICON_INFORMATION)

