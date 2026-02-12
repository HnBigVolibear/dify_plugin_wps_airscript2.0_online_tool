
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_add_worksheet(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        sheet_name = tool_parameters.get("sheet_name", None)
        newSheetName = "创建失败！"
        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))
            result = client.add_worksheet(sheet_name.strip() if sheet_name and sheet_name.strip() else None) # type: ignore
            if result and result[0]['success']:
                newSheetName = result[0]['sheetName']
            elif result and result[0].get('error', "") == 'CoreExecError: E_NAME_CONFLICT':
                raise Exception("工作表名称已存在，请重新输入工作表名称！！！")
            else:
                raise Exception(f"添加工作表失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"添加工作表失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message(f"添加工作表成功！新建的工作表名称是：{newSheetName if sheet_name else '自动命名：'+newSheetName}\n")
