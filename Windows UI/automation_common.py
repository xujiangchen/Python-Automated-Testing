import time
from typing import Dict, Any, Optional, Tuple, List
import uiautomation as automation


class AutomationCommon:

    def find_element(self, location: Dict[str, Any], parent: Optional[automation.Control] = None, max_wait: float = 10, interval: float = 0.5, retry=True):
        """
        找到指定元素
        :param location: 元素的定位信息 {"control_type": "WindowControl", "searchDepth":1, "Name":"Sign In", "ClassName":"", .......}
        :param parent: 该元素的父元素
        :param max_wait: 最大查找时间
        :param interval: 每次重试间隔时间
        :param retry: 是否要使用兜底方法 （通过部分定位参数，查找父元素 30 深度的所有元素，尝试是否能够找到）
        """
        # 如果没有传入parent，则以windows桌面作为父层级
        if parent is None:
            parent = automation.GetRootControl()
        if "control_type" not in location:
            raise ValueError("定位必须包含 'control_type' 属性 ")
        control_type = location["control_type"]
        if not hasattr(parent, control_type):
            raise ValueError(f"无效的控件类型{control_type}")
        # 获取查询深度，如果没有默认深度为 5
        search_depth = location.pop("searchDepth", 5)
        ctrl_method = getattr(parent, control_type)
        # 生成匹配属性数据
        match_kwargs = {k: v for k, v in location.items() if k != 'control_type'}
        start_time = time.time()
        while time.time() - start_time <= max_wait:
            try:
                ctrl = ctrl_method(seatchDepth=search_depth, **match_kwargs)
                # 存在就返回，windows UI 的树中存在这个元素
                if ctrl.Exists(0.1):
                    return ctrl
            except Exception as e:
                print(f"Retrying due to: {e}")
            time.sleep(interval)
        # 如果在指定的深度没有找到对应的组件，则通过其他属性在windows的树中查找是否还有符合这个属性的组件
        if retry:
            for _ in range(3):
                index, ele = self.get_control_index(parent=parent, location=location)
                if ele is not None:
                    return ele
                time.sleep(5)
        raise AssertionError(f"[Failed] Cannot find: {control_type} with {match_kwargs} at depth {search_depth}")

    def wait_element_visible(self, location: Dict[str, Any], parent: Optional[automation.Control] = None, max_wait: float = 10, interval: float = 0.5):
        """
        等待这个元素可见
        :param location: 元素的定位信息 {"control_type": "WindowControl", "searchDepth":1, "Name":"Sign In", "ClassName":"", .......}
        :param parent: 该元素的父元素
        :param max_wait: 最大查找时间
        :param interval: 每次重试间隔时间
        """
        for _ in range(3):
            try:
                # 首先找到这个元素
                ctrl = self.find_element(parent=parent, location=location, max_wait=max_wait, interval=interval)
                # 判断是否被隐藏，不可见，不在屏幕内
                if not ctrl.IsOffscreen:
                    return ctrl
            except Exception as e:
                print(f"Retry due to: {e}")
        raise AssertionError(f"[Failed]: {location} 不可见！")

    def click_element(self, location: Dict[str, Any], parent: Optional[automation.Control] = None, max_wait: float = 10, interval: float = 0.5):
        """
        点击元素
        :param location: 元素的定位信息 {"control_type": "WindowControl", "searchDepth":1, "Name":"Sign In", "ClassName":"", .......}
        :param parent: 该元素的父元素
        :param max_wait: 最大查找时间
        :param interval: 每次重试间隔时间
        """
        for _ in range(3):
            try:
                # 首先判断这个元素是否可见
                ctrl = self.wait_element_visible(parent=parent, location=location, max_wait=max_wait, interval=interval)
                time.sleep(0.5)
                # 这个元素是否可以被操作
                if not ctrl.IsEnabled:
                    continue
                try:
                    ctrl.Click()
                    return True
                except Exception as e:
                    print(f"点击报错: {e}, 尝试使用 GetClickablePoint 进行点击")
                    #  将鼠标移动到元素的位置，再次尝试点击
                    x, y, clickable = ctrl.GetClickablePoint()
                    if clickable:
                        automation.MoveTo(x, y)
                        automation.Click(x, y)
                        return True
            except Exception as e:
                print(f"Retry due to: {e}")
            time.sleep(interval)
            raise RuntimeError(f"[failed] : {location} 不能被点击!")
        raise AssertionError("")

    def fill_element(self, text, location: Dict[str, Any], parent: Optional[automation.Control] = None, max_wait: float = 10, interval: float = 0.5):
        """
        向文本输入框中输入数据
        :param text: 所要输入的文本信息
        :param location: 元素的定位信息 {"control_type": "WindowControl", "searchDepth":1, "Name":"Sign In", "ClassName":"", .......}
        :param parent: 该元素的父元素
        :param max_wait: 最大查找时间
        :param interval: 每次重试间隔时间
        """
        for _ in range(3):
            try:
                ctrl = self.wait_element_visible(parent=parent, location=location, max_wait=max_wait, interval=interval)
                if not ctrl.IsEnabled:
                    continue
                try:
                    ctrl.Click()
                except Exception as e:
                    print(f"元素点击失败：{e}")
                automation.SendKeys('{Ctrl}a{Del}')
                time.sleep(0.3)
                ctrl.SendKeys(text, waitTime=1)
                return True
            except Exception as e:
                print(f"[Exception] 输入文本失败：{e}")
                time.sleep(interval)
        raise AssertionError(f"[Failed]: 元素 {location} 不能被输入！")

    @staticmethod
    def get_control_index(parent: automation.Control, location: Dict[str, Any], search_depth: int = 30, equals_match: bool = False) \
            -> Tuple[Optional[int], Optional[automation.Control]]:
        """
        根据父元素查找指定属性的元素所在层级
        :param parent:
        :param location:
        :param search_depth:
        :param equals_match:
        :return:
        """
        # 参数校验
        if parent is None:
            raise ValueError("父元素不能为空！")
        if "control_type" not in location:
            raise ValueError("元素类型不能为空！")

        # 获取组件类型
        control_type = location['control_type']
        # 获取所有用于匹配组件的属性
        match_fields = {k: v for k, v in location.items() if k not in ("control_type", "searchDepth")}

        # 尝试在所有同类型的元素中进行匹配
        stack: List[Tuple[automation.Control, int]] = [(parent, 0)]
        while stack:
            node, depth = stack.pop()
            if depth >= search_depth:
                continue
            for child in node.GetChildren():
                child_depth = depth + 1
                if child.ControlTypeName == control_type:
                    matched = True
                    for attr, expected in match_fields.items():
                        actual = getattr(child, attr, None)
                        if actual is None:
                            matched = False
                            break
                        if equals_match:
                            if actual != expected:
                                matched = False
                                break
                        else:
                            if str(expected) not in str(actual):
                                matched = False
                                break
                    if matched:
                        return child_depth, child
                if child_depth < search_depth:
                    stack.append((child, child_depth))
        return None, None
