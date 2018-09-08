/* eslint-disable */
var idTmr;
var getExplorer = () => {
    let explorer = window.navigator.userAgent;
    //ie
    if (explorer.indexOf("MSIE") >= 0 || explorer.indexOf("rv:11") >= 0) {
        return 'ie';
    }
    //firefox

    else if (explorer.indexOf("Firefox") >= 0) {
        return 'Firefox';
    }
    //Chrome
    else if (explorer.indexOf("Chrome") >= 0) {
        return 'Chrome';
    }
    //Opera
    else if (explorer.indexOf("Opera") >= 0) {
        return 'Opera';
    }
    //Safari
    else if (explorer.indexOf("Safari") >= 0) {
        return 'Safari';
    }
}
// 判断浏览器是否为IE
var exportToExcel = (data, name) => {

    // 判断是否为IE
    if (getExplorer() == 'ie') {
        tableToIE(data, name)
    } else {
        tableToNotIE(data, name)
    }
}

var Cleanup = () => {
    window.clearInterval(idTmr);
}

// ie浏览器下执行
var tableToIE = (data, name) => {
    var curTbl = data;
    var oXL = new ActiveXObject("Excel.Application");

    //创建AX对象excel
    var oWB = oXL.Workbooks.Add();
    //获取workbook对象
    var xlsheet = oWB.Worksheets(1);
    //激活当前sheet
    var sel = document.body.createTextRange();
    
    var tableContainer = document.getElementById('j-table2export')
    
    tableContainer.innerHTML = curTbl;
    sel.moveToElementText(document.getElementById('j-table2export'));
    //sel.moveToElementText(curTbl);
    //把表格中的内容移到TextRange中
    sel.select();
    //全选TextRange中内容
    sel.execCommand("Copy");
    //复制TextRange中内容   
    xlsheet.Paste();
    //粘贴到活动的EXCEL中

    oXL.Visible = true;
    //设置excel可见属性

    // let fname;
    // try {
    //     fname = oXL.Application.GetSaveAsFilename(name +".xls", "Excel Spreadsheets (*.xls), *.xls");
    // } catch (e) {
    //     print("Nested catch caught " + e);
    // } finally {
    //     let savechanges;

    //     oWB.SaveAs(fname);

    //     oWB.Close(savechanges=false);
    //     //xls.visible = false;
    //     oXL.Quit();
    //     oXL = null;
    //     // 结束excel进程，退出完成
    //     // window.setInterval("Cleanup();", 1);
    //     // idTmr = window.setInterval("Cleanup();", 1);
    //     // Cleanup();
    // }
}

// 非ie浏览器下执行
var tableToNotIE = (function() {
    // 编码要用utf-8不然默认gbk会出现中文乱码
    var uri = 'data:application/vnd.ms-excel;base64,',
        template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>',
        base64 = function(s) {
            return window.btoa(unescape(encodeURIComponent(s)));

        },

        format = (s, c) => {
            return s.replace(/{(\w+)}/g,
                (m, p) => {
                    return c[p];
                })
        }
    return (table, name) => {
        var ctx = {
            worksheet: name,
            table
        }

        //创建下载
        var link = document.createElement('a');
        link.setAttribute('href', uri + base64(format(template, ctx)));

        link.setAttribute('download', name);


        // window.location.href = uri + base64(format(template, ctx))
        link.click();
    }
})()

// 导出函数
var export2Excel = (theadData, tbodyData, dataname) => {
    // let re = /http/ // 字符串中包含http,则默认为图片地址
    var th_len = theadData.length // 表头的长度
    var tb_len = tbodyData.length // 记录条数
    var width = 60 // 设置图片大小
    var height = 90

    // 添加表头信息
    var thead = '<thead><tr>'
    for (var i = 0; i < th_len; i++) {
        thead += '<th>' + theadData[i] + '</th>'
    }
    thead += '</tr></thead>'

    // 添加每一行数据
    var tbody = '<tbody>'
    for (var i = 0; i < tb_len; i++) {
        tbody += '<tr>'
        var row = tbodyData[i] // 获取每一行数据

        for (var key in row) {
            if (Array.isArray(row[key])) { // 如果为数组则视为图片数组，则需要加div包住图片
                var imgArray = [...row[key]];
                var imgContent = '';
                imgArray.forEach((item, index) => {
                    imgContent += '<img src=\'' + item + '\' ' + 'width=' + '\"' + width + '\"' + ' ' + 'height=' + '\"' + height + '\"' + '>'
                    tbody += '<td><div>' + imgContent + '</div></td>'
                })
            } else {
                tbody += '<td style="text-align:center; width:200; height:50;">' + row[key] + '</td>'
            }
        }
        tbody += '</tr>'
    }
    tbody += '</tbody>'

    var table = '<table>' + thead + tbody + '</table>'
    // 导出表格
    exportToExcel(table, dataname)
}

// export default export2Excel;