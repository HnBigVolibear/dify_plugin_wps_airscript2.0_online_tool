
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_set_font(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")
        font_name = tool_parameters.get("font_name", None)
        font_size = tool_parameters.get("font_size", None)
        bold = tool_parameters.get("bold", None)
        italic = tool_parameters.get("italic", None)
        color = tool_parameters.get("color", None)
        if color:
            color = color.strip().lower()
        if color and color[0] != "#":
            color = "#" + color
        if not is_hex_color(color):
            raise Exception("请填入正确的颜色值，例如：#FF0000")
        if font_size and font_size < 1:
            raise Exception("请填入正确的字号，例如：12。注：字号不能小于1！")


        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的区域地址，例如：A1:D3")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))

            font_options = {}
            if font_name:
                font_options["name"] = font_name
            if font_size:
                font_options["size"] = font_size
            if bold:
                font_options["bold"] = bold=="是"
            if italic:
                font_options["italic"] = italic=="是"
            if color:
                # color = client.rgb_to_excel_color(color_r, color_g, color_b)
                font_options["color"] = color

            result = client.set_font(address, font_options, sheet_name)
            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"设置字体失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"设置字体失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"设置字体成功！区域：{address}\n")

def is_hex_color(s):
    import re
    """判断字符串是否为标准十六进制色值（已小写化）"""
    return bool(re.match(r'^#([0-9a-f]{3}|[0-9a-f]{6}|[0-9a-f]{8})$', s))