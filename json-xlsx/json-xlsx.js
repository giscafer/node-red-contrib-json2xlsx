const xlsx = require('json-as-xlsx');

module.exports = function (RED) {
  function jsonToXlsx(config) {
    RED.nodes.createNode(this, config);
    const node = this;
    node.on('input', function (msg) {
      const sheetName = config.sheetName || msg.sheetName || 'Sheet1';
      const name = config.name || msg.name;
      msg.filename = `${
        config.filename || msg.filename || name || 'temp'
      }.xlsx`;
      msg.description = config.description || msg.description || 'excel file';
      const content = ToXlsx(JsonToTable(msg.payload), sheetName, msg.filename);
      msg.payload = content;
      msg.topic = `${config.topic || msg.topic || name} from Node-RED`;
      msg.content = content;
      node.send(msg);
    });
  }

  function JsonToTable(payload) {
    const columns = [];
    const content = [];

    if (Array.isArray(payload)) {
      if (payload.length > 0) {
        let c = 0;
        for (const p in payload[0]) {
          columns.push({ label: p, value: `col${c++}` });
        }
        payload.forEach((row) => {
          const itemRow = {};
          let c = 0;
          for (const p in row) {
            itemRow[`col${c++}`] = row[p];
          }
          content.push(itemRow);
        });
      }
    }
    return { columns, content };
  }

  function ToXlsx(table, sheetName, fileName, download = false) {
    const settings = {
      fileName,
      extraLength: 3, // A bigger number means that columns will be wider
      writeMode: download ? 'writeFile' : 'write', // The available parameters are 'WriteFile' and 'write'. This setting is optional. Useful in such cases https://docs.sheetjs.com/docs/solutions/output#example-remote-file
      writeOptions: { bookType: 'xlsx', type: 'buffer' }, // Style options from https://docs.sheetjs.com/docs/api/write-options
      // RTL: true, // Display the columns from right-to-left (the default value is false)
    };
    table.sheet = sheetName;
    return xlsx([table], settings); // Will download the excel file
  }
  RED.nodes.registerType('json-xlsx', jsonToXlsx);
};
