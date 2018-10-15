import FileSaver from 'file-saver';
import XLSX from 'xlsx';
/**
 * @description ie9以上表格下载方法
 * @param domId 表格id
 * @param fileName 文件名称
 */
export function exportExcel(domId, fileName) {
  if (
    navigator.appName == 'Microsoft Internet Explorer' &&
    navigator.appVersion.match(/9./i) == '9.'
  ) {
    //ie9中下载excel 整个表格拷贝到EXCEL中
    var curTbl = document.querySelector('#' + domId);
    var oXL;
    try {
      //创建AX对象excel
      oXL = new ActiveXObject('Excel.Application');
    } catch (e) {
      alert(
        '要下载该表，您必须安装Excel电子表格软件，同时浏览器须使用“ActiveX 控件”，您的浏览器须允许执行控件。'
      );
      return null;
    }
    //获取workbook对象
    var oWB = oXL.Workbooks.Add();
    //激活当前sheet
    var oSheet = oWB.ActiveSheet;
    var sel = document.body.createTextRange();
    //把表格中的内容移到TextRange中
    console.log(sel);
    sel.moveToElementText(curTbl);
    //全选TextRange中内容
    sel.select();
    //复制TextRange中内容
    sel.execCommand('Copy');
    //粘贴到活动的EXCEL中
    oSheet.Paste();
    //设置excel可见属性
    oXL.Visible = true;
    var fname = oXL.Application.GetSaveAsFilename(
      fileName + '.xlsx',
      'Excel Spreadsheets (*.xlsx), *.xlsx'
    );
    oWB.SaveAs(fname);
    oWB.Close();
    oXL.Quit();
  } else {
    //ie9以上
    var wb = XLSX.utils.table_to_book(document.querySelector('#' + domId));
    var wbout = XLSX.write(wb, {
      bookType: 'xlsx',
      bookSST: true,
      type: 'array'
    });
    try {
      FileSaver.saveAs(
        new Blob([wbout], { type: 'application/octet-stream' }),
        fileName + '.xlsx'
      );
    } catch (e) {
      if (typeof console !== 'undefined') {
        console.log(e, wbout);
      }
    }
    return wbout;
  }
}

export default exportExcel;
