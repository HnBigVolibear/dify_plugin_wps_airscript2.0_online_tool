from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_rename_worksheet(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        old_sheet_name = tool_parameters.get("old_sheet_name", "")
        new_sheet_name = tool_parameters.get("new_sheet_name", "")

        if not old_sheet_name:
            raise Exception("请填入原工作表名称")
        if not new_sheet_name:
            raise Exception("请填入新工作表名称")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.rename_worksheet(old_sheet_name=old_sheet_name, new_sheet_name=new_sheet_name)
            if result and result[0].get('success'):
                pass
            else:
                raise Exception(f"重命名工作表失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"重命名工作表失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"工作表 '{old_sheet_name}' 已成功重命名为 '{new_sheet_name}'\n")
