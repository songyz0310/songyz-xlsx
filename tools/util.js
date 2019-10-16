import XLSX from '../xlsx';
import XLSX1 from 'xlsx';
import {
  bodyStyle,
  titleStyle
} from './common';
import {
  getCharCol,
  writeFile
} from './support';


//导入文件的类型
export const xlsxTypes = ["xlsx", "xlc", "xlm", "xls", "xlt", "xlw", "csv"];

/**
 * 导入文件
 * @param { File } file 
 * @param { Object } opts 
 */
export const importSlsx = (file, opts) => {
  return new Promise(function (resolve, reject) {
    const reader = new FileReader()
    reader.onload = function (e) {
      opts = opts || {};

      opts.type = 'binary';
      opts._dateType = opts._dateType || 1; //1,"yyyy-MM-dd hh:mm",2,时间戳
      opts._numberType = opts._numberType || 1; //1,不使用科学计数法,2,使用科学计数法

      const wb = XLSX.read(e.target.result, opts);
      resolve(Object.keys(wb.Sheets).map(key => XLSX.utils.sheet_to_json(wb.Sheets[key])).reduce((prev, next) => prev.concat(next)))
    }
    reader.readAsBinaryString(file.raw)
  })
}

/**
 * 导出数据
 * @param { Array<Object> } dataArray 
 * @param { String } fileName 
 * @param { String } sheetName 
 */
export const exportXlsx = (dataArray, fileName, sheetName) => {
  let type = 'xlsx';
  dataArray = dataArray || [{}];
  fileName = fileName || 'file';
  sheetName = sheetName || 'mySheet';

  var keyMap = Object.keys(dataArray[0]);
  var title = {};
  keyMap.forEach(key => title[key] = key);
  dataArray.unshift(title);

  //用来保存转换好的json 
  var sheetData = [];
  // console.log(JSON.stringify(dataArray));
  dataArray.map((row, i) => {
    let style = i == 0 ? titleStyle : bodyStyle;
    return keyMap.map((key, j) => {
      return {
        style: style,
        value: row[key],
        position: (j > 25 ? getCharCol(j) : String.fromCharCode(65 + j)) + (i + 1)
      };
    })
  }).reduce((prev, next) => prev.concat(next)).forEach((cell, i) =>
    sheetData[cell.position] = {
      v: cell.value,
      s: cell.style
    }
  );
  var outputPos = Object.keys(sheetData); //设置区域,比如表格从A1到D10

  var wb = {
    SheetNames: [sheetName], //保存的表标题
    Sheets: {}
  };

  wb.Sheets[sheetName] = Object.assign({}, sheetData, //内容
    {
      '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1] //设置填充区域
    }
  );

  wb.Sheets[sheetName]['!cols'] = keyMap.map(key => {
    let col = {};
    col.wpx = key.length <= 4 ? 70 : 90;
    return col;
  });


  var buffer = XLSX.write(wb, {
    bookType: type,
    bookSST: false,
    type: 'buffer'
  });

  writeFile(fileName + "." + type, buffer);
}

/**
 * 导出数据:包含一个条件的数据表格
 * @param { Array<Object> } dataArray 
 * @param { String } condition 
 * @param { String } fileName 
 * @param { String } sheetName 
 */
export const exportXlsxWithCondition = (dataArray, condition, fileName, sheetName) => {
  if (condition) {
    exportXlsxWithConditions(dataArray, [condition], fileName, sheetName);
  } else {
    exportXlsx(dataArray, fileName, sheetName);
  }
}

/**
 * 导出数据:包含多个条件的数据表格
 * @param { Array<Object> } dataArray 
 * @param { Array<String> } conditions 
 * @param { String } fileName 
 * @param { String } sheetName 
 */
export const exportXlsxWithConditions = (dataArray, conditions, fileName, sheetName) => {
  debugger
  let type = 'xlsx';
  dataArray = dataArray || [{}];
  fileName = fileName || 'file';
  conditions = conditions || [""];
  sheetName = sheetName || 'mySheet';

  var keyMap = Object.keys(dataArray[0]);
  var title = {};
  keyMap.forEach(key => title[key] = key);
  dataArray.unshift(title);

  dataArray.unshift(...conditions)

  var ws = XLSX1.utils.aoa_to_sheet(dataArray.map((row, i) => {
    if (i < conditions.length) {
      return [conditions[i]];
    } else {
      return keyMap.map((key, j) => {
        return row[key];
      })
    }
  }));

  ws['!merges'] = conditions.map((c, i) => {
    return {
      s: {
        r: i,
        c: 0
      },
      e: {
        r: i,
        c: keyMap.length - 1
      }
    };
  });;

  for (const key in ws) {
    if (key.indexOf("!") < 0) {
      const cell = ws[key];
      if (parseInt(key.replace(/\D+/, "")) <= conditions.length) {
        cell.s = titleStyle;
      } else {
        cell.s = bodyStyle;
      }
    }
  }

  var wb = XLSX1.utils.book_new();
  XLSX1.utils.book_append_sheet(wb, ws, "sheet1");

  const buffer = XLSX.write(wb, {
    type: "buffer",
    bookType: "xlsx"
  });

  writeFile(fileName + "." + type, buffer);
}
