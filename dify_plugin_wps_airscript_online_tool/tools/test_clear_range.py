
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_clear_range(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        address = tool_parameters.get("address", "").strip()
        sheet_name = tool_parameters.get("sheet_name", "")
        # 新增参数：是否保留格式
        keep_format = tool_parameters.get("keep_format", "否") == "是"

        if not address or address[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
            raise Exception("请填入正确的区域地址，例如：A1:D3")
        
        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            if keep_format:
                result = client.clear_range_contents(address, sheet_name)
                if result and result[0].get('success'):
                    pass
                else:
                    raise Exception(f"清除区域内容失败！WPS官方返回错误信息：{result}")
            else:
                result = client.clear_range(address, sheet_name)
                if result and result[0].get('success') and result[0].get('message') == "清除成功":
                    pass
                else:
                    raise Exception(f"清除区域失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"清除区域失败！WPS官方返回错误信息：{e}")

        if keep_format:
            yield self.create_text_message("清除区域内容成功（已保留格式）\n")
        else:
            yield self.create_text_message("清除区域成功！（内容和格式均已清除）\n")
