import os
import xlrd
import qrcode


def readExcel(file_path):
    # 读取Excel文件
    excel_data_obj = xlrd.open_workbook(file_path)

    # 读取指定 sheet名，可修改
    this_excel_sheet_name = excel_data_obj.sheet_names()[0]
    this_excel_data = excel_data_obj.sheet_by_name(this_excel_sheet_name)

    data_rows = this_excel_data.nrows
    data_list = []
    for data_row in range(data_rows):
        this_row = this_excel_data.row_values(data_row)
        this_row = str(this_row[0])
        # 过滤空单元格
        if len(this_row) > 0 and this_row[0] !=' ':
            this_row = eval(this_row)
            # int型在读取后会转换为float，需要还原
            if type(this_row) == float:
                this_row_div = this_row / int(this_row)
                if this_row_div == 1:
                    this_row = int(this_row)
            data_list.append(str(this_row))

    return data_list



def drowQRcode(project_name, data_list):
    # 生产目标文件目录
    isExists = os.path.exists('./' + project_name)
    if not isExists:
        os.mkdir('./' + project_name)

    for data_row in data_list:
        img_file_path = os.path.join(project_name, data_row + '.png')

        # 实例化QRCode生成qr对象
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=6,
            border=2
        )
        # 传入数据
        qr.add_data(data_row)
        qr.make(fit=True)
        # 生成二维码
        img = qr.make_image()
        # 保存二维码
        img.save(img_file_path)




if __name__ == '__main__':
    excel_path = input(r'请输入Excel文件路径：')
    project_name = input(r'请输入项目名称：')
    data_list = readExcel(file_path='./test.xlsx')
    drowQRcode(project_name=project_name, data_list=data_list)
