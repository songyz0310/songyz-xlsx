<p align="center">
  <img width="320" src="https://avatars1.githubusercontent.com/u/32382526?s=460&v=4">
</p>


[English](./README.md) | 简体中文

## 简介

[songyz-xlsx](https://github.com/songyz0310/songyz-xlsx) 是一个处理Excel的工具，它基于 [js-xlsx](https://github.com/SheetJS/js-xlsx) 进行的扩展开发，借以实现公司部分业务需求，如有问题，可以在 [关于js-xlsx的使用](https://www.cnblogs.com/songyz/p/10345469.html) 进行留言评论，方便大家共同学习进步。

## 功能

```
- 导入时间的处理

- 导入关于数字的处理
  - 科学计数法

```

## 开发

```bash
# 使用npm 安装依赖
npm install songyz-style --save

# 核心代码
# import example
import XLSX from 'songyz-xlsx'

export const importSlsx = (file, opts) => {
    return new Promise(function (resolve, reject) {
        const reader = new FileReader()
        reader.onload = function (e) {
            opts = opts || {};

            opts.type = 'binary';
            opts._dateType = opts._dateType || 1; //1,"yyyy-MM-dd hh:mm",2,时间戳
            opts._numberType = opts._numberType || 1; //1,不适用科学计数法,2,使用科学计数法

            const wb = XLSX.read(e.target.result, opts);
            resolve(Object.keys(wb.Sheets).map(key => XLSX.utils.sheet_to_json(wb.Sheets[key])).reduce((prev, next) => prev.concat(next)))
        }
        reader.readAsBinaryString(file.raw)
    })
}

# export example 
export const exportXlsx = (dataArray, fileName) => {
    let type = 'xlsx';
    dataArray = dataArray || [{}];
    fileName = fileName || 'file';

    var keyMap = Object.keys(dataArray[0]);
    var title = {};
    keyMap.forEach(key => title[key] = key);
    dataArray.unshift(title);

    var sheetData = [];

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
    var outputPos = Object.keys(sheetData); 

    var wb = {
        SheetNames: ['mySheet'], 
        Sheets: {
            'mySheet': Object.assign({},
                sheetData, 
                {
                    '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1] 
                }
            )
        }
    };
    
    var buffer = XLSX.write(wb, { bookType: type, bookSST: false, type: 'buffer' });

    writeFile(fileName + "." + type, buffer);
}

```

## Donate

如果你觉得这个项目帮助到了你，你可以帮作者买一杯果汁表示鼓励 :tropical_drink:
![donate](./dist/pay.jpg)


## License

[MIT](./dist/LICENSE)

Copyright (c) 2019-present songyz
