### xlsxParser ([English Document](../../README.md))
> Convert complex excel into the json you want


### introduce
* Currently only.xlsx format is supported。
* This project is based on nodejs. Please install the nodejs environment first。
* Execute the command
```bash
# Clone this repository
git clone https://github.com/sliencio/xlsxParser.git
# Go into the repository
cd xlsxParser
# Install dependencies
npm install
```

* 配置config.json
```javascript
{
  "xlsx": {
    //Type in line
    "typeRow":2,
    //Line of title
    "headRow": 3,
    //xlsx path
    "src": "./excel/**/[^~$]*.xlsx",
    //output path
    "outPath": "./json",
    //array separator
    "arrayFlag":",",
    //The join symbol of a subtable
    "sonCon": ".",
    //Sibling table join symbol
    "brotherCon": "&"
  },
  "json": {
    //is compression json
    "compress": false
  }
}
```
### run

#### linux or unix

``` shell
  bash export.sh
```

#### window

​	Double-click on export. Bat

### export file

​	Read `. / excel / ` all XLSX file path, export to path `. / json / `, json file named after the name sheet

#### Example 1 basic functionality (see resources./excel/basic.xlsx)   
![excel](./docs/image/excel.png)

The output is as follows (item.extro :) : (sibling table: item&brother) {array subtable item.attrs}

```json
{
  "1001": {
    "id": "1001",
    "Des": "这是机枪",
    "name": "机枪",
    "attack": 150,
    "expAdd": 10,
    "extro": {
      "id": "1001",
      "extro": "机枪的额外属性"
    },
    "attrs": [
      {
        "id": "1001",
        "value": 10,
        "name": "名称1"
      },
      {
        "id": "1001",
        "value": 15,
        "name": "名称2"
      },
      {
        "id": "1001",
        "value": 20,
        "name": "名称3"
      },
      {
        "id": "1001",
        "value": 25,
        "name": "名称4"
      },
      {
        "id": "1001",
        "value": 30,
        "name": "名称5"
      },
      {
        "id": "1001",
        "value": 35,
        "name": "名称6"
      }
    ],
    "life": 200
  }
}
```



### The following data types are supported

* ` number ` numeric types.
* ` Boolean ` Boolean.
* ` string ` string.
* ` date ` date type.
* ` object ` object, consistent with JS object.
* ` array ` array, consistent with JS array.
* ` id ` primary key type (when the table has id type, as to the hash json format output, or output) in an array format.
* ` id [] ` primary key array, is only found in from the table.



### Rule header

* when basic data types (string,number,bool) are set, they are generally not automatically judged, but data types can also be explicitly declared.
* string type: named form ` column name string `.
* numeric types: named form ` column number `.
* date type: ` column date `. The date format should conform to the standard date format. Such as ` YYYY/M/D H: M: s ` or ` ` YYYY/M/D, and so on.
Boolean type: * named form ` column bool `.
* array: named form ` column names [] `.
* object: named form ` column {} `.
* primary key: named form ` column id `, there can be only one column in the table.
* : the primary key array named form ` column id [] `, there can be only one column in the table, only exist in from the table.
* column name to `! ` beginning do not export.



### sheet rule

- the sheet name to `! ` beginning do not export the table.
- table from the table name ` name `. ` child table ` ` the main table name ` & table ` ` brother, the main table must be in the front of the child table or brothers tables.

### notes

* the key symbols are all English half Angle symbols, which are consistent with JSON requirements.
* object writing is the same as object writing in JavaScript
* arrays are written the same way as arrays are written in JavaScript

### supplement

* run on Windows/MAC/Linux.
* project address [xlsxParser] (https://github.com/sliencio/xlsxParser)
