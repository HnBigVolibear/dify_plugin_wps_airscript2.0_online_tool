
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_set_row_height(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        row_index = tool_parameters.get("row_index", 1)
        height = tool_parameters.get("height", 20)
        sheet_name = tool_parameters.get("sheet_name", "")

        if row_index < 1:
            raise Exception("行索引必须大于等于1")

        if height < 1:
            raise Exception("行高必须大于等于1")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.set_row_height(row_index, height, sheet_name)
            if result and result[0]['success']:
                pass
            else:
                raise Exception(f"设置行高失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"设置行高失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"设置行高成功！第{row_index}行的高度设置为{height}磅\n")
