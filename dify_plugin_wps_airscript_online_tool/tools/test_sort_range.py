
from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from wps_airscript_client import WPSAirScriptClient

class test_sort_range(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        sortList = tool_parameters.get("sortList", "").strip()
        modeHeader = tool_parameters.get("modeHeader", "xlGuess")
        modeMatchCase = tool_parameters.get("modeMatchCase", "否")
        # has_header = tool_parameters.get("has_header", True)
        sheet_name = tool_parameters.get("sheet_name", "")
        if not sheet_name:
            raise Exception("请填入正确的sheet名称")

        import json
        try:
            sortList = json.loads(sortList)
        except:
            try:
                import ast
                sortList = ast.literal_eval(sortList)
            except Exception as e:
                raise Exception('请填入正确的排序规则！例如：[ ["C", "desc"], ["A", "asc"] ]')
        
        if not sortList or not isinstance(sortList, list):
            raise Exception('请填入正确的排序规则！例如：[ ["C", "desc"], ["A", "asc"] ]')
        for col in sortList:
            if len(col) != 2:
                raise Exception('请填入正确的排序规则！例如：[ ["C", "desc"], ["A", "asc"] ]')
            if col[1] not in ["asc", "desc"]:
                raise Exception('请填入正确的排序规则！例如：[ ["C", "desc"], ["A", "asc"] ]')
            if col[0] not in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
                raise Exception('请填入正确的排序规则！例如：[ ["C", "desc"], ["A", "asc"] ]')

        try:
            file_id = self.session.storage.get("file_id").decode('utf-8')
            token = self.session.storage.get("token").decode('utf-8')
            script_id = self.session.storage.get("script_id").decode('utf-8')

            client = WPSAirScriptClient(file_id=file_id, token=token, script_id=script_id ,base_url=self.session.storage.get("base_url").decode('utf-8'))

            sortOptions = {
                "modeHeader": modeHeader, # 是否有表头参与，xlGuess为自动，xlYes为包含表头，xlNo为不包含表头。默认设置为xlGuess
                "modeMatchCase": modeMatchCase, # 是否大小写敏感，“是”为区分大小写，“否”为不区分大小写，默认设置“否”
            }

            # result = client.sort_range(address, sort_options, sheet_name)
            # sortList = [ ["C", "desc"], ["A", "asc"] ]
            result = client.sortUsedRange( sheet_name=sheet_name, sortList=sortList , sortOptions=sortOptions)
            if result:
                pass
            else:
                raise Exception(f"排序失败！WPS官方返回错误信息：{result}")
        except Exception as e:
            raise Exception(f"排序失败！WPS官方返回错误信息：{e}")

        yield self.create_text_message("排序成功！\n")
