
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_set_alignment(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")
        horizontal = tool_parameters.get("horizontal", None)
        vertical = tool_parameters.get("vertical", None)
        if horizontal:
            horizontal = int(horizontal)
        if vertical:
            vertical = int(vertical)
        if not horizontal and not vertical:
            raise Exception("闹呢！？水平和垂直方向都不选，那你用俺这个命令搞锤子啊？请至少设置一个对齐方式！")

        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的区域地址，例如：A1:D3")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))

            align_options = {}
            if horizontal:
                align_options["horizontal"] = horizontal
            if vertical:
                align_options["vertical"] = vertical

            result = client.set_alignment(address, align_options, sheet_name)
            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"设置对齐方式失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"设置对齐方式失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"设置对齐方式成功！区域：{address}\n")
