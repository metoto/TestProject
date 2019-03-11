# -*- coding: utf-8 -*-
import docx
import os
import GoogleTranslateUtils
import ExcelUtils
import WordUtils



class test:

    def __init__(self):
        pass

if __name__ == "__main__":
    path=u"E:\\_测试组文件\\测试资源\\机翻测试.xlsx"
    sheetName = u'word文档翻译原文生成'
    tempPath = u"E:\\_00000\\TemplateWord.docx"
    destFolder_path = u"E:\\_00000\\TestDoc\\"





    workbook = ExcelUtils.ExcelDriver(path)
    worksheet = workbook.getWorksheet(sheetName)

    text1 = ExcelUtils.ExcelDriver.getCellData(2, 4, worksheet)
    text2 = ExcelUtils.ExcelDriver.getCellData(2, 5, worksheet)
    text3 = ExcelUtils.ExcelDriver.getCellData(2, 6, worksheet)
    text4 = ExcelUtils.ExcelDriver.getCellData(2, 7, worksheet)
    print text1

    js=GoogleTranslateUtils.Py4Js()
    tk1 = js.getTk(text1)
    tk2 = js.getTk(text2)
    tk3 = js.getTk(text3)
    tk4 = js.getTk(text4)

    num = workbook.getSheetRowCount(worksheet)
    for index in range(num-1):
        langId = workbook.getCellData(index+3,3,worksheet)
        print langId
        transText1 = GoogleTranslateUtils.translate(text1, tk1, langId)
        transText2 = GoogleTranslateUtils.translate(text2, tk2, langId)
        transText3 = GoogleTranslateUtils.translate(text3, tk3, langId)
        transText4 = GoogleTranslateUtils.translate(text4, tk4, langId)
        print transText1
        resultCell1 = worksheet.cell(index + 3, 4)
        resultCell2 = worksheet.cell(index + 3, 5)
        resultCell3 = worksheet.cell(index + 3, 6)
        resultCell4 = worksheet.cell(index + 3, 7)

        workbook.setCellData(worksheet, resultCell1, transText1)
        workbook.setCellData(worksheet, resultCell2, transText2)
        workbook.setCellData(worksheet, resultCell3, transText3)
        workbook.setCellData(worksheet, resultCell4, transText4)

    workbook.saveExcelFile(path)
    #print TempCell

    for index in range(num-1):
        langId = workbook.getCellData(index + 3, 3, worksheet)
        caseId = workbook.getCellData(index + 3, 1, worksheet)

        paraStr1 = workbook.getCellData(index + 3, 4, worksheet)
        paraStr2 = workbook.getCellData(index + 3, 5, worksheet)
        paraStr3 = workbook.getCellData(index + 3, 6, worksheet)
        paraStr4 = workbook.getCellData(index + 3, 7, worksheet)

        dest_path = destFolder_path+str(caseId)+'_'+langId+'.docx'
        context = {'paragraph1': paraStr1, 'list1': paraStr2,'list2': paraStr3, 'list3':paraStr4 }
        WordUtils.DocUtils.applyTempleDoc(tempPath, dest_path, context)
    #paraStr = workbook.getCellData(index + 3, 3, worksheet)
    #context = {'paragraph1': paraStr }
    # print os.path.exists(path)
    # print os.path.exists(dest_path)
    # pathlist = os.path.split(dest_path)
    # folderPath = pathlist[0]
    # fileName = pathlist[1]
    # print folderPath
    # print fileName

    # file_tmp = docx.Document(path)
    #
    # #file_tmp = docx.Document()
    # file_tmp.add_paragraph(u'第一段')
    # file_tmp.add_paragraph(u'第二段')
    #
    # file_tmp.save(dest_path)