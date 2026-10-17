process.argv[2] = __dirname + "/_tmp.docx";
const o2 = require("./anexo2.js"); const o3 = require("./anexo3.js");
require("fs").writeFileSync(__dirname + "/obs.json", JSON.stringify({ OS: o2.OBS, OE: o3.OE, OM: o3.OM }));
