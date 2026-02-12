
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_delete_worksheet(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        sheet_identifier = tool_parameters.get("sheet_identifier", "")

        if not sheet_identifier:
            raise Exception("请输入工作表名称！！！")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.delete_worksheet(sheet_identifier)
            if result and result[0]['success']:
                yield self.create_text_message("删除工作表成功！\n")
            else:
                # yield self.create_text_message("删除工作表失败！\n")
                raise Exception("删除工作表失败！可能是你要删除的表名不存在？")
        except Exception as e:
            raise Exception(f"删除工作表失败！WPS官方返回错误信息：{e}")
