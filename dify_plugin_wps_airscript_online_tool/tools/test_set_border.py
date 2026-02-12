
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_set_border(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")
        line_style = tool_parameters.get("line_style", None)
        weight = tool_parameters.get("weight", None)
        color = tool_parameters.get("color", "#3B3B3B")
        if color:
            color = color.strip().lower()
        if color and color[0] != "#":
            color = "#" + color
        if not is_hex_color(color):
            raise Exception("请填入正确的颜色值，例如：#FF0000")
        if not line_style and not weight:
            raise Exception("闹呢！？边框线样式和边框粗细都不选，那你用俺这个命令搞锤子啊？请至少设置一个边框属性！")

        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的区域地址，例如：A1:D3")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))

            border_options = {}
            if line_style:
                border_options["lineStyle"] = int(line_style)
            if weight:
                border_options["weight"] = int(weight)

            if color:
                # border_options["color"] = client.rgb_to_excel_color(color_r, color_g, color_b)
                border_options["color"] = color

            result = client.set_border(address, border_options, sheet_name)
            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"设置边框失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"设置边框失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"设置边框成功！区域：{address}\n")


def is_hex_color(s):
    import re
    """判断字符串是否为标准十六进制色值（已小写化）"""
    return bool(re.match(r'^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$', s))