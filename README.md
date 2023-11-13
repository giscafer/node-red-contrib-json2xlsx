# node-red-contrib-json2xlsx

base on: [json-as-xlsx](https://github.com/LuisEnMarroquin/json-as-xlsx)

`json-xlsx` node convert an object array to an Xlsx buffered file output.
For example you could convert an sql table output to a file, o send it as attachment by e-mail.

## Usage

### Attributes

```html
<p>Converts the payload from a json object array into a buffered xlsx file</p>
<h3>Details</h3>
<p><code>msg.payload</code> should contains a JSON array of objects.</p>
<p>
  <code>File name</code>
  文件名称，可以使用<code>msg.name</code>来动态设置文件名称
</p>
<p>
  <code>Sheet name</code> Sheet
  名称，可以使用<code>msg.sheetName</code>来动态设置Sheet名称
</p>
<p><code>Description</code> 描述信息，发送邮件时的邮件内容</p>
```

## License

MIT
