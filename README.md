# node-red-contrib-json2xlsx

base on: [json-as-xlsx](https://github.com/LuisEnMarroquin/json-as-xlsx)

`json-xlsx` node convert an object array to an Xlsx buffered file output.
For example you could convert an sql table output to a file, o send it as attachment by e-mail.

## Usage

### Example flow

```json
[
  {
    "id": "25dbe814222163d1",
    "type": "tab",
    "label": "流程 1",
    "disabled": false,
    "info": "",
    "env": []
  },
  {
    "id": "9d3f77b57829d00d",
    "type": "function",
    "z": "25dbe814222163d1",
    "name": "Excel数据",
    "func": "msg.payload = [\n    { name: '哈哈', age: 16 },\n    { name: '哈哈1', age: 16 },\n]\nreturn msg;",
    "outputs": 1,
    "timeout": 0,
    "noerr": 0,
    "initialize": "",
    "finalize": "",
    "libs": [],
    "x": 300,
    "y": 200,
    "wires": [["1a8cd83c7f5690f3"]]
  },
  {
    "id": "db7d5a096ee0bdb7",
    "type": "inject",
    "z": "25dbe814222163d1",
    "name": "",
    "props": [],
    "repeat": "",
    "crontab": "",
    "once": false,
    "onceDelay": 0.1,
    "topic": "",
    "x": 170,
    "y": 200,
    "wires": [["9d3f77b57829d00d"]]
  },
  {
    "id": "4220bfbef7af0259",
    "type": "e-mail",
    "z": "25dbe814222163d1",
    "server": "smtp.163.com",
    "port": "465",
    "authtype": "BASIC",
    "saslformat": true,
    "token": "oauth2Response.access_token",
    "secure": true,
    "tls": true,
    "name": "test@gmail.com",
    "dname": "",
    "x": 680,
    "y": 200,
    "wires": []
  },
  {
    "id": "4d514166bebd7a3e",
    "type": "debug",
    "z": "25dbe814222163d1",
    "name": "debug 108",
    "active": true,
    "tosidebar": true,
    "console": false,
    "tostatus": false,
    "complete": "true",
    "targetType": "full",
    "statusVal": "",
    "statusType": "auto",
    "x": 670,
    "y": 160,
    "wires": []
  },
  {
    "id": "1a8cd83c7f5690f3",
    "type": "json-xlsx",
    "z": "25dbe814222163d1",
    "name": "",
    "filename": "",
    "topic": "",
    "sheetName": "Sheet1",
    "description": "",
    "x": 490,
    "y": 200,
    "wires": [["4220bfbef7af0259", "4d514166bebd7a3e"]]
  }
]
```

## License

MIT
