# ExcelToJson

用于读取Excel文件将内容转化为Json

## 用法
```c#
    var excelToJson = new ExcelToJson();
    string json = excelToJson.OpenExcelAndToJson("../../../Res/Examples/ExampleData.xlsx");
```

## 示例Excel文件
| Id  | Name   | Birthday   | Sex   | Hobbies               | LuckyNumbers | OtherInfo                                                             |
|-----|--------|------------|-------|-----------------------|--------------|-----------------------------------------------------------------------|
| int | string | DateTime   | bool  | string[]              | int[]        | object                                                                |
| id  | 姓名     | 出生日期       | 性别    | 爱好                    | 幸运数字         | 其他信息                                                                  |
| 1   | 张三     | 1987/6/5   | TRUE  | ["Singing","Dancing"] | [1,3,5,7,9]  | {}                                                                    |
| 2   | 李四     | 2012/12/20 | FALSE | ["a","b"]             | [2,4,6,8,10] | {"Address":"China","Info":{},"Age":9,"NUll":null,"Nest":{"Nest1":{}}} |
| 3   | 王五     | 2022/2/22  | FALSE | ["c","ddd"]           | [0]          |

## 转为Json的结果

```js
[
  {
    "Id": 1,
    "Name": "张三",
    "Birthday": "1987-06-05T00:00:00",
    "Sex": true,
    "Hobbies": [
      "Singing",
      "Dancing"
    ],
    "LuckyNumbers": [
      1,
      3,
      5,
      7,
      9
    ],
    "OtherInfo": {}
  },
  {
    "Id": 2,
    "Name": "李四",
    "Birthday": "2012-12-20T00:00:00",
    "Sex": false,
    "Hobbies": [
      "a",
      "b"
    ],
    "LuckyNumbers": [
      2,
      4,
      6,
      8,
      10
    ],
    "OtherInfo": {
      "Address": "China",
      "Info": {},
      "Age": 9,
      "NUll": null,
      "Nest": {
        "Nest1": {}
      }
    }
  },
  {
    "Id": 3,
    "Name": "王五",
    "Birthday": "2022-02-22T00:00:00",
    "Sex": false,
    "Hobbies": [
      "c",
      "ddd"
    ],
    "LuckyNumbers": [
      0
    ],
    "OtherInfo": null
  }
]
```
