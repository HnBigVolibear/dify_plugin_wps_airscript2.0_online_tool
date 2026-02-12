from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_delete_all_pivot_tables(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        sheet_name = tool_parameters.get("sheet_name", "")

        if not sheet_name:
            raise Exception("请填入工作表名称")

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.delete_all_pivot_tables(sheet_name=sheet_name)
            if result and result[0].get('success'):
                pass
            else:
                raise Exception(f"删除透视表失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"删除透视表失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(str(result[0].get('message', '透视表删除成功'))+"\n")
