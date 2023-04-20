import { Workbook } from 'exceljs'

const obj = {
  billList: [
    { BillDate: '2023-04-18 12:49:35', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 134.44, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-03-29 11:07:29', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 1, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-03-24 08:31:34', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 412, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-03-23 13:01:20', BillType: '办公用品', GroupName: 'BDMAP房地一体软件开发', Money: 29, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-03-15 09:15:02', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 158.91, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-02-20 18:59:22', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 149.6, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-02-15 16:46:21', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 67.5, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-02-15 10:39:14', BillType: '加油费', GroupName: 'BDMAP房地一体软件开发', Money: 100, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-02-07 17:30:51', BillType: '交通费', GroupName: 'BDMAP房地一体软件开发', Money: 24, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-02-04 16:00:55', BillType: '交通费', GroupName: 'BDMAP房地一体软件开发', Money: 19.67, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-02-04 14:19:33', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 14, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-02-01 14:10:10', BillType: '办公用品', GroupName: 'BDMAP房地一体软件开发', Money: 30, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-01-31 16:17:31', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 159.19, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-01-13 20:33:12', BillType: '加油费', GroupName: 'BDMAP房地一体软件开发', Money: 100, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-01-13 08:23:07', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 690, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-01-13 08:22:17', BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', Money: 2174, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' },
    { BillDate: '2023-01-08 16:23:55', BillType: '办公费其他', GroupName: '综合部日常办公', Money: 506.84, Name: '测试', Remarks: '部门加班零食', address: '河南郑州' }
  ],
  amount: [
    { BillType: '办公用品', GroupName: 'BDMAP房地一体软件开发', totalMoney: 59 },
    { BillType: '餐饮费', GroupName: 'BDMAP房地一体软件开发', totalMoney: 3960.64 },
    { BillType: '加油费', GroupName: 'BDMAP房地一体软件开发', totalMoney: 200 },
    { BillType: '交通费', GroupName: 'BDMAP房地一体软件开发', totalMoney: 43.67 },
    { BillType: '办公费其他', GroupName: '综合部日常办公', totalMoney: 506.84 }
  ]
}

// 导出excel
function exportExcel() {
  // https://github.com/exceljs/exceljs/blob/HEAD/README_zh.md
  if (!obj.billList.length && !obj.amount.length) return ElMessage('请先查询再导出')

  const workbook = new Workbook() // 创建工作簿
  const worksheet = workbook.addWorksheet('Sheet1') // 添加工作表

  // 插入前三行
  worksheet.insertRow(1, [])
  worksheet.insertRow(2, [])
  worksheet.insertRow(3, [])

  // 合并单元格
  worksheet.mergeCells('A1:I2')
  worksheet.mergeCells('G3:I3')
  worksheet.getCell('G3').value = '日期：' + formatTime()

  // 插入表头
  worksheet.insertRow(4, ['序号', '时间', '群组', '用途', '消费金额', '有票金额', '备注', '消费地址', '账单提交人'])

  let totalMoney = 0 // 合计总额
  const headers = ['BillDate', 'GroupName', 'BillType', 'Money', 'Money', 'Remarks', 'address', 'Name']
  obj.billList.forEach((element, index) => {
    totalMoney = (totalMoney * 100 + element.Money * 100) / 100
    const temp = []
    temp.push(index + 1)
    headers.forEach(item => {
      if (item === 'Money') temp.push(element[item].toFixed(2))
      else temp.push(element[item])
    })

    worksheet.addRow(temp) // 添加行
  })

  // 插入表格下方的合计行
  worksheet.insertRow(3 + 1 + obj.billList.length + 1, ['合计', '', '', '', totalMoney.toFixed(2), totalMoney.toFixed(2), '', '', ''])

  const group = ['序号', '账单类型', '总额'],
    type = []

  // 取出群组和类型
  obj.amount.forEach(element => {
    if (group.indexOf(element.GroupName) === -1) group.push(element.GroupName)
    if (type.indexOf(element.BillType) === -1) type.push(element.BillType)
  })

  // 创建二维数组
  const twoArray = []

  // 行
  twoArray[0] = group // 第一行赋值表头
  for (let i = 1; i < type.length + 1 + 1; i++) {
    twoArray[i] = []
    // 列
    for (let j = 2; j < group.length; j++) {
      twoArray[i][0] = i // 第一列赋值索引
      if (type[i - 1]) twoArray[i][1] = type[i - 1] // 第二列赋值类型
      else twoArray[i][1] = '' // 第二列赋值类型 (类型为空 赋值为空)
      twoArray[i][j] = '0.00' // 其余金额单元格赋值为0
    }
  }
  twoArray[type.length + 1][0] = '合计' // 最后一行赋值为合计 */

  // 给每个符合条件的分别赋值金额
  for (let i = 0; i < obj.amount.length; i++) {
    let row = 0,
      col = 0
    for (let j = 3; j < group.length; j++) {
      if (obj.amount[i].GroupName == group[j]) col = j
    }
    for (let x = 0; x < type.length; x++) {
      if (obj.amount[i].BillType == type[x]) row = x
    }
    twoArray[row + 1][col] = obj.amount[i].totalMoney.toFixed(2)
  }

  // 计算出合计总数 每一行第三列 最后一行第二列往后
  for (let i = 1; i < twoArray.length - 1; i++) {
    for (let j = 3; j < twoArray[i].length; j++) {
      twoArray[i][2] = ((twoArray[i][2] * 100 + twoArray[i][j] * 100) / 100).toFixed(2)
      twoArray[twoArray.length - 1][j] = ((twoArray[twoArray.length - 1][j] * 100 + twoArray[i][j] * 100) / 100).toFixed(2)
      if (i === twoArray.length - 1 - 1) {
        twoArray[twoArray.length - 1][2] = ((twoArray[twoArray.length - 1][2] * 100 + twoArray[twoArray.length - 1][j] * 100) / 100).toFixed(2)
      }
    }
  }

  // 添加二维数组的每一行
  for (let i = 0; i < twoArray.length; i++) worksheet.addRow(twoArray[i])

  worksheet.addRow([]) // 往后添加一行空的
  // 添加最后的签名成员行
  worksheet.addRow(['经手人：', '', '', '复核：', '', '财务：', '', '审核：', ''])

  // 根据每一列的名字设置对应的列宽
  worksheet.columns = [
    { header: '序号', width: 12 },
    { header: '时间', width: 18.11 },
    { header: '群组', width: 25 },
    { header: '用途', width: 26.44 },
    { header: '消费金额', width: 24 },
    { header: '有票金额', width: 24 },
    { header: '备注', width: 28.78 },
    { header: '消费地址', width: 25 },
    { header: '账单提交人', width: 26.33 }
  ]

  // 遍历工作表中具有值的所有行
  worksheet.eachRow(function (row, rowNumber) {
    // 设置导出数据行高
    if (rowNumber > 4 && rowNumber < 3 + 1 + obj.billList.length + 1) row.height = 60

    // 设置字体和文字对齐
    row.alignment = {
      wrapText: true,
      vertical: 'middle',
      horizontal: 'center'
    }
    row.font = {
      name: '宋体',
      size: 14
    }

    // 连续遍历所有非空单元格   并添加边框
    row.eachCell(function (cell, colNumber) {
      if (rowNumber > 3 && rowNumber < 3 + 1 + obj.billList.length + 1 + twoArray.length + 1) {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        }
      }
    })

    // 规定行进行加粗
    if (rowNumber === 4) row.font.bold = true
    if (rowNumber === 3 + 1 + obj.billList.length + 1) row.font.bold = true
    if (rowNumber === 3 + 1 + obj.billList.length + 1 + 1) row.font.bold = true
    if (rowNumber === 3 + 1 + obj.billList.length + 1 + twoArray.length) row.font.bold = true
    if (rowNumber === 3 + 1 + obj.billList.length + 1 + twoArray.length + 2) row.font.bold = true
  })

  // 遍历此列中的所有当前单元格
  worksheet.getColumn('H').eachCell(function (cell, rowNumber) {
    // 设置H列 消费地址列样式 缩小字体
    if (rowNumber > 4 && rowNumber < 3 + 1 + obj.billList.length + 1) {
      cell.font = {
        name: '宋体',
        size: 12
      }
      cell.alignment = {
        wrapText: true,
        vertical: 'middle',
        horizontal: 'left'
      }
    }
  })
  // 设置第一列样式 加粗
  worksheet.getColumn('A').eachCell(function (cell, rowNumber) {
    if (rowNumber > 3) {
      cell.font = {
        name: '宋体',
        bold: true,
        size: 14
      }
    }
  })

  worksheet.getCell('A1').value = '报销辅助单'
  worksheet.getCell('A1').font.size = 20
  worksheet.getCell('A1').font.bold = true
  worksheet.getCell('G3').font.bold = true

  workbook.xlsx.writeBuffer().then(data => {
    const blob = new Blob([data], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    })
    openDownloadDialog(blob, formatTime() + '帐单列表')
  })
}

function openDownloadDialog(url, saveName) {
  if (typeof url == 'object' && url instanceof Blob) {
    url = URL.createObjectURL(url) // 创建blob地址
  }
  const aLink = document.createElement('a')
  aLink.href = url
  aLink.download = saveName || '' // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
  let event
  if (window.MouseEvent) event = new MouseEvent('click')
  else {
    event = document.createEvent('MouseEvents')
    event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null)
  }
  aLink.dispatchEvent(event)
}

function formatTime() {
  const date = new Date(),
    yaer = date.getFullYear(),
    month = date.getMonth() + 1,
    day = date.getDate()
  return yaer + '年' + month + '月' + day + '日'
}
