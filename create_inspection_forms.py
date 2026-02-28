"""
隧道工程20项检验批验收用表Excel生成脚本
"""
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

def create_full_excel():
    wb = Workbook()

    # 定义样式
    header_font = Font(name='宋体', size=14, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    sub_header_font = Font(name='宋体', size=11, bold=True)
    sub_header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 定义20个分项的验收内容 - 区分主控项目和一般项目
    # 主控项目带△标记
    inspection_items = {
        "1-隧道总体": {
            "name": "隧道总体",
            "basic": [
                "隧道衬砌内轮廓及所有运营设施均不得侵入建筑限界",
                "洞口设置应满足设计要求",
                "洞内外的排水系统设置应满足设计要求",
                "高速公路、一级公路和二级公路隧道拱部、边墙、路面、设备箱洞应不渗水"
            ],
            "main_control": [
                ("行车道宽度(mm)", "±10", "尺量：曲线每20m、直线每40m检查1个断面"),
                ("内轮廓宽度(mm)", "不小于设计值", "激光测距仪：曲线每20m、直线每40m检查1个断面"),
                ("内轮廓高度(mm)△", "不小于设计值", "全站仪：曲线每20m、直线每40m检查1个断面"),
                ("隧道偏位(mm)", "20", "全站仪：曲线每20m、直线每40m测1处")
            ],
            "general": [
                ("边坡或仰坡坡度", "不大于设计值", "尺量：每洞口检查10处")
            ],
            "general_check": [
                "洞口边坡、仰坡应无落石",
                "排水系统应不淤积、不堵塞"
            ]
        },
        "2-明洞浇筑": {
            "name": "明洞浇筑",
            "basic": [
                "基础的地基承载力应满足设计要求并符合施工技术规范规定，严禁超挖后回填虚土",
                "钢筋的加工及安装应满足设计要求",
                "明洞与暗洞连接应满足设计要求"
            ],
            "main_control": [
                ("混凝土强度(MPa)△", "在合格标准内", "按附录D检查"),
                ("混凝土厚度(mm)△", "不小于设计值", "尺量：每10m检查1个断面")
            ],
            "general": [
                ("墙面平整度(mm)", "施工缝、变形缝处20；其他部位5", "2m直尺：每10m每侧连续检查2尺")
            ],
            "general_check": [
                "蜂窝麻面面积不得超过该面总面积的0.5%，深度不得超过10mm",
                "隧道衬砌钢筋混凝土结构裂缝宽度不得超过0.2mm"
            ]
        },
        "3-明洞防水层": {
            "name": "明洞防水层",
            "basic": [
                "防水层铺设前应清除喷射混凝土表面的外露钢筋头",
                "防水层应无空鼓、折皱、破损等缺陷",
                "防水层拼接应满足设计要求"
            ],
            "main_control": [
                ("防水层完整性△", "无破损", "尺量：每5m检查1处"),
                ("搭接宽度(mm)△", "符合设计，且不小于100", "尺量：每个搭接处检查1点")
            ],
            "general": [
                ("防水层与基面密贴", "符合设计", "尺量：每5m检查1处")
            ],
            "general_check": []
        },
        "4-明洞回填": {
            "name": "明洞回填",
            "basic": [
                "明洞回填土的压实度应满足设计要求",
                "明洞拱背回填应对称进行"
            ],
            "main_control": [
                ("回填土压实度△", "符合设计", "按附录B检查：每100m²检查1个点")
            ],
            "general": [
                ("回填厚度(mm)", "符合设计", "尺量：每20m检查1处"),
                ("边坡坡度", "符合设计", "尺量：每20m检查1处")
            ],
            "general_check": ["回填坡面应不积水"]
        },
        "5-洞身开挖": {
            "name": "洞身开挖",
            "basic": [
                "当围岩自稳能力差时，开挖前应做好预加固、预支护",
                "应采用控制爆破减少开挖对围岩的扰动",
                "应严格控制欠挖，拱脚、墙脚以上1m范围内严禁欠挖",
                "洞身开挖在清除浮石后应及时进行初喷支护"
            ],
            "main_control": [
                ("拱部超挖(mm)△", "Ⅰ级围岩平均100最大200；ⅡⅢⅣ级平均150最大250；ⅤⅥ级平均100最大150", "全站仪：每20m检查1个断面"),
                ("边墙超挖(mm)△", "每侧+100,0；全宽+200,0", "尺量：每20m检查1个断面"),
                ("仰拱、隧底超挖(mm)△", "平均100，最大250", "水准仪：每20m检查3处")
            ],
            "general": [],
            "general_check": ["洞顶应无浮石"]
        },
        "6-喷射混凝土": {
            "name": "喷射混凝土",
            "basic": [
                "喷射混凝土的强度应满足设计要求",
                "喷射混凝土应分层喷射",
                "喷射混凝土表面应无裂缝"
            ],
            "main_control": [
                ("喷射混凝土强度(MPa)△", "在合格标准内", "试件检验：每10m³或每工班不少于1组"),
                ("喷层厚度(mm)△", "平均≥设计厚度；最小≥0.5设计厚度且≥50", "凿孔法或雷达仪：每10m检查1个断面")
            ],
            "general": [
                ("喷层与围岩粘结力(MPa)", "符合设计", "拉拔试验：每10m²抽查1处")
            ],
            "general_check": ["喷层应无裂缝、脱落"]
        },
        "7-锚杆": {
            "name": "锚杆",
            "basic": [
                "锚杆的材质、类型、规格、数量、性能应满足设计要求",
                "锚杆安装位置应满足设计要求",
                "锚杆的锚固质量应满足设计要求"
            ],
            "main_control": [
                ("锚杆长度(m)△", "不小于设计值", "尺量或孔深仪：每20m检查5根"),
                ("锚杆间距(mm)△", "±100", "尺量：每20m检查5根"),
                ("锚杆方向(°)△", "符合设计", "测量：每20m检查5根"),
                ("锚杆拔力(kN)△", "≥设计值", "拔力试验：每100根不少于1组")
            ],
            "general": [],
            "general_check": ["锚杆垫板与岩面间应无间隙"]
        },
        "8-钢筋网": {
            "name": "钢筋网",
            "basic": ["钢筋网铺设应在初喷混凝土后进行"],
            "main_control": [
                ("钢筋网保护层厚度(mm)△", "≥20", "凿孔法：每10m测5点"),
                ("网格尺寸(mm)△", "±10", "尺量：每100m²检查3个网眼")
            ],
            "general": [
                ("搭接长度(mm)", "≥50", "尺量：每20m测3点")
            ],
            "general_check": ["钢筋网与锚杆或其他固定构件连接不得松脱"]
        },
        "9-钢架": {
            "name": "钢架",
            "basic": [
                "钢架之间应采用纵向钢筋连接，安装基础应牢固",
                "钢架应紧靠初喷面",
                "连接钢板与钢架应焊接牢固"
            ],
            "main_control": [
                ("榀数(榀)△", "不少于设计值", "目测或尺量：逐榀检查"),
                ("间距(mm)△", "±50", "尺量：逐榀检查"),
                ("垂直度(°)△", "≤2", "测量：逐榀检查")
            ],
            "general": [
                ("保护层厚度(mm)", "≥25", "测量：逐榀检查")
            ],
            "general_check": []
        },
        "10-仰拱": {
            "name": "仰拱",
            "basic": [
                "仰拱的开挖长度应满足设计要求",
                "仰拱的底板高程应满足设计要求",
                "仰拱的钢筋安装应满足设计要求"
            ],
            "main_control": [
                ("混凝土强度(MPa)△", "在合格标准内", "试件检验：每工班不少于1组"),
                ("仰拱厚度(mm)△", "不小于设计值", "尺量：每20m检查1个断面")
            ],
            "general": [
                ("仰拱底板高程(mm)", "±50", "水准仪：每20m检查1处")
            ],
            "general_check": []
        },
        "11-仰拱回填": {
            "name": "仰拱回填",
            "basic": [
                "仰拱回填材料的强度应满足设计要求",
                "仰拱回填的压实度应满足设计要求"
            ],
            "main_control": [
                ("回填压实度△", "符合设计", "核子密度仪或灌砂法：每50m²检查1点")
            ],
            "general": [
                ("回填厚度(mm)", "符合设计", "尺量：每20m检查1处")
            ],
            "general_check": []
        },
        "12-衬砌钢筋": {
            "name": "衬砌钢筋",
            "basic": [
                "钢筋的品种、级别、规格、数量应满足设计要求",
                "钢筋的连接方式应满足设计要求",
                "钢筋的安装位置应满足设计要求"
            ],
            "main_control": [
                ("主筋间距(mm)△", "±10", "尺量：每20m检查2个断面"),
                ("箍筋间距(mm)△", "±20", "尺量：每20m检查5个间距"),
                ("两层钢筋间距(mm)△", "±10", "尺量：每20m检查3处"),
                ("钢筋保护层厚度(mm)△", "＋10，－5", "尺量：每20m检查3处")
            ],
            "general": [],
            "general_check": []
        },
        "13-混凝土衬砌": {
            "name": "混凝土衬砌",
            "basic": [
                "混凝土的强度应满足设计要求",
                "混凝土的抗渗等级应满足设计要求",
                "混凝土的厚度应满足设计要求"
            ],
            "main_control": [
                ("混凝土强度(MPa)△", "在合格标准内", "试件检验：每工班不少于1组"),
                ("混凝土抗渗等级△", "不小于设计值", "抗渗试验：按规范频率"),
                ("衬砌厚度(mm)△", "不小于设计值", "雷达仪或取芯：每40m检查1个断面")
            ],
            "general": [
                ("表面平整度(mm)", "20", "2m靠尺：每40m检查2处")
            ],
            "general_check": [
                "蜂窝麻面面积不得超过该面总面积的0.5%",
                "裂缝宽度不得超过0.2mm"
            ]
        },
        "14-防水层": {
            "name": "防水层",
            "basic": [
                "防水层铺设前应清除喷射混凝土表面的外露钢筋头",
                "防水层应无空鼓、折皱、破损等缺陷"
            ],
            "main_control": [
                ("防水层完整性△", "无破损", "充气试验：每5环检查1处"),
                ("搭接宽度(mm)△", "≥100（双焊缝）", "尺量：每个搭接处检查1点")
            ],
            "general": [
                ("搭接缝质量", "饱满、连续", "目测：全数检查")
            ],
            "general_check": []
        },
        "15-止水带": {
            "name": "止水带",
            "basic": [
                "止水带的材质、规格应满足设计要求",
                "止水带的安装位置应满足设计要求",
                "止水带的接头质量应满足设计要求"
            ],
            "main_control": [
                ("止水带宽度(mm)△", "符合设计", "尺量：每环检查1处"),
                ("止水带厚度(mm)△", "符合设计", "尺量：每环检查1处"),
                ("止水带埋设位置(mm)△", "±30", "尺量：每环检查3处")
            ],
            "general": [],
            "general_check": []
        },
        "16-排水": {
            "name": "排水",
            "basic": [
                "排水管道的材质、规格应满足设计要求",
                "排水管道的安装应满足设计要求",
                "排水管道应畅通"
            ],
            "main_control": [
                ("排水管位置△", "符合设计", "尺量：每20m检查1处"),
                ("排水管坡度(%)△", "≥2，且符合设计", "水准仪：每20m检查1处"),
                ("排水管畅通性△", "畅通", "通水试验：每50m检查1处")
            ],
            "general": [],
            "general_check": []
        },
        "17-超前锚杆": {
            "name": "超前锚杆",
            "basic": [
                "超前锚杆的材质、规格应满足设计要求",
                "超前锚杆的安装应满足设计要求",
                "超前锚杆的注浆质量应满足设计要求"
            ],
            "main_control": [
                ("超前锚杆长度(m)△", "不小于设计值", "尺量：每循环检查5根"),
                ("超前锚杆间距(mm)△", "±100", "尺量：每循环检查5根"),
                ("超前锚杆方向(°)△", "符合设计", "测量：每循环检查5根")
            ],
            "general": [],
            "general_check": []
        },
        "18-超前小导管": {
            "name": "超前小导管",
            "basic": [
                "超前小导管的材质、规格应满足设计要求",
                "超前小导管的安装应满足设计要求",
                "超前小导管的注浆质量应满足设计要求"
            ],
            "main_control": [
                ("小导管长度(m)△", "不小于设计值", "尺量：每循环检查10根"),
                ("小导管间距(mm)△", "±50", "尺量：每循环检查10根"),
                ("小导管方向(°)△", "符合设计", "测量：每循环检查10根")
            ],
            "general": [],
            "general_check": []
        },
        "19-管棚": {
            "name": "管棚",
            "basic": [
                "管棚的材质、规格应满足设计要求",
                "管棚的安装应满足设计要求",
                "管棚的注浆质量应满足设计要求"
            ],
            "main_control": [
                ("管棚长度(m)△", "不小于设计值", "尺量：每根检查"),
                ("管棚间距(mm)△", "±50", "尺量：每5根检查1根"),
                ("钢管规格△", "符合设计", "尺量：每根检查"),
                ("布设角度(°)△", "符合设计", "测量：每根检查")
            ],
            "general": [],
            "general_check": []
        },
        "20-路面": {
            "name": "隧道路面",
            "basic": [
                "基层的强度、厚度应满足设计要求",
                "面层的强度、厚度，平整度应满足设计要求"
            ],
            "main_control": [
                ("基层/面层压实度△", "≥设计值", "核子密度仪或灌砂法：每200m²1处"),
                ("基层/面层厚度(mm)△", "±10", "尺量：每200m测1处"),
                ("弯沉值△", "符合设计", "弯沉仪：每车道每20m测1处")
            ],
            "general": [
                ("平整度(mm)", "基层≤8，面层≤5", "3m直尺：每200m测2处×10尺")
            ],
            "general_check": []
        }
    }

    # 创建工作表
    for key, data in inspection_items.items():
        ws = wb.create_sheet(title=data["name"])

        # 标题
        ws.merge_cells('A1:G1')
        ws['A1'] = f'检验批质量验收记录 - {data["name"]}'
        ws['A1'].font = Font(name='宋体', size=16, bold=True)
        ws['A1'].alignment = center_align

        # 基本信息
        row = 3
        info_items = [
            ('工程名称：', 'A'), ('隧道名称：', 'B'), ('里程范围：', 'C'),
            ('分项工程：', 'D'), ('检验批编号：', 'E'), ('验收日期：', 'F')
        ]
        for label, col in info_items:
            ws[f'{col}{row}'] = label
            ws[f'{col}{row}'].font = Font(name='宋体', size=11)
            ws[f'{chr(ord(col)+1)}{row}'] = ''
            ws[f'{chr(ord(col)+1)}{row}'].border = Border(bottom=Side(style='thin'))

        # 基本要求
        if data["basic"]:
            row += 1
            ws.merge_cells(f'A{row}:G{row}')
            ws[f'A{row}'] = '一、基本要求'
            ws[f'A{row}'].font = sub_header_font
            ws[f'A{row}'].fill = sub_header_fill

            for item in data["basic"]:
                row += 1
                ws[f'A{row}'] = item
                ws.merge_cells(f'A{row}:G{row}')
                ws[f'A{row}'].alignment = left_align

        # 主控项目
        if data.get("main_control"):
            row += 1
            ws.merge_cells(f'A{row}:G{row}')
            ws[f'A{row}'] = '二、主控项目'
            ws[f'A{row}'].font = sub_header_font
            ws[f'A{row}'].fill = sub_header_fill

            row += 1
            headers = ['项次', '检查项目', '规定值或允许偏差', '实测值', '偏差', '检验方法', '检验频率']
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=row, column=col, value=h)
                cell.font = sub_header_font
                cell.fill = sub_header_fill
                cell.border = thin_border
                cell.alignment = center_align

            for i, item in enumerate(data["main_control"]):
                row += 1
                ws.cell(row=row, column=1, value=i+1).border = thin_border
                ws.cell(row=row, column=2, value=item[0]).border = thin_border
                ws.cell(row=row, column=3, value=item[1]).border = thin_border
                ws.cell(row=row, column=4, value='').border = thin_border
                ws.cell(row=row, column=5, value='').border = thin_border
                ws.cell(row=row, column=6, value=item[2]).border = thin_border
                ws.cell(row=row, column=7, value='').border = thin_border
                for col in range(1, 8):
                    ws.cell(row=row, column=col).alignment = center_align if col < 6 else left_align

        # 一般项目
        if data.get("general"):
            row += 1
            ws.merge_cells(f'A{row}:G{row}')
            ws[f'A{row}'] = '三、一般项目'
            ws[f'A{row}'].font = sub_header_font
            ws[f'A{row}'].fill = sub_header_fill

            row += 1
            headers = ['项次', '检查项目', '规定值或允许偏差', '实测值', '偏差', '检验方法', '检验频率']
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=row, column=col, value=h)
                cell.font = sub_header_font
                cell.fill = sub_header_fill
                cell.border = thin_border
                cell.alignment = center_align

            for i, item in enumerate(data["general"]):
                row += 1
                ws.cell(row=row, column=1, value=i+1).border = thin_border
                ws.cell(row=row, column=2, value=item[0]).border = thin_border
                ws.cell(row=row, column=3, value=item[1]).border = thin_border
                ws.cell(row=row, column=4, value='').border = thin_border
                ws.cell(row=row, column=5, value='').border = thin_border
                ws.cell(row=row, column=6, value=item[2]).border = thin_border
                ws.cell(row=row, column=7, value='').border = thin_border
                for col in range(1, 8):
                    ws.cell(row=row, column=col).alignment = center_align if col < 6 else left_align

        # 外观质量
        if data.get("general_check"):
            row += 1
            ws.merge_cells(f'A{row}:G{row}')
            ws[f'A{row}'] = '四、外观质量'
            ws[f'A{row}'].font = sub_header_font
            ws[f'A{row}'].fill = sub_header_fill

            for item in data["general_check"]:
                row += 1
                ws[f'A{row}'] = item
                ws.merge_cells(f'A{row}:G{row}')
                ws[f'A{row}'].alignment = left_align

        # 验收签字
        row += 2
        ws[f'A{row}'] = '施工单位检查评定：'
        ws[f'B{row}'] = '□合格  □不合格'
        ws[f'E{row}'] = '质检员：'

        row += 1
        ws[f'A{row}'] = '监理单位验收意见：'
        ws[f'B{row}'] = '□同意验收  □不同意验收'
        ws[f'E{row}'] = '监理工程师：'

        # 设置列宽
        for col in range(1, 8):
            ws.column_dimensions[get_column_letter(col)].width = 16

    # 删除默认sheet
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # 保存
    output_path = 'E:/AI-Work/Lobster01/隧道工程验收用表/隧道工程20项检验批验收用表.xlsx'
    wb.save(output_path)
    print(f'Excel文件创建成功！保存至: {output_path}')

if __name__ == '__main__':
    create_full_excel()
