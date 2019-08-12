import excel_helper
import pretty_json

json = excel_helper.open_excel_json("file\新建 Microsoft Excel 工作表.xlsx", remove_enter=True)

print(pretty_json.format_json(json))
