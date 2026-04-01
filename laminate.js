const express = require("express");
const app = express();
const PORT = process.env.PORT || 3000;
const session = require("express-session");

app.use(session({
  secret: "icl-secret",
  resave: false,
  saveUninitialized: true
}));
const ExcelJS = require("exceljs");   // KEEP THIS ONE
const axios = require("axios");       // ADD THIS ONE
app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));
app.use(express.json()); 

// =======================
// STUDY → KIT MAP
// =======================

const studyToKitsMap = {
  "ICL-CABNATe-FAMCRU": [
    "Infant Cohort 3 Screening",
    "Infant Cohort 3 Visit 2,4,6 Day 3,16,42",
    "Infant Cohort 3 Visit 3 Day 9",
    "Infant Cohort 3 Visit 5,EW Day 28"
  ]
};

// =======================
// EQUIPMENT DEFINITIONS
// =======================

const infantPack = {
  name: "Infant Phlebotomy Pack 1",
  img: "/images/infant_phlebotomy_pack.jpg",
  qty: 1,
  desc: "Plaster, Webcol, cotton wool, 21G Butterfly and Bulldog",
  inst: "No Label, No Barcodes"
};

const edta05 = {
  name: "0.5ml EDTA Microcontainer",
  
  img: "/images/edta_0_5ml.jpg",
  qty: 1,
  desc: "EDTA Microtainer",
  inst: "Label applied in a manner that allows a clear vertical view of the tube contents. IC Labs CT Barcode"
};

const edta08 = {
  name: "0.8ml EDTA Microcontainer",
 
  img: "/images/edta_0_8ml.jpg",
  qty: 1,
  desc: "EDTA Microtainer",
  inst: "Label applied in a manner that allows a clear vertical view of the tube contents. IC Labs CT Barcode"
};

const sst05 = {
  name: "0.5ml SST Microcontainer",
  
  img: "/images/sst_0_5ml.jpg",
  qty: 1,
  desc: "SST Microtainer",
  inst: "Label applied in a manner that allows a clear vertical view of the tube contents. IC Labs CT Barcode"
};

const sst08 = {
  name: "0.8ml SST Microcontainer",
  img: "/images/sst_0_8ml.jpg",
  qty: 1,
  desc: "SST Microtainer",
  inst: "Label applied in a manner that allows a clear vertical view of the tube contents. IC Labs CT Barcode"
};

// =======================
// KIT TEMPLATES
// =======================

const kitTemplates = {
  "Infant Cohort 3 Screening": [
    infantPack, edta05, sst08, edta08, sst05
  ],
  "Infant Cohort 3 Visit 2,4,6 Day 3,16,42": [
    infantPack, sst05
  ],
  "Infant Cohort 3 Visit 3 Day 9": [
    infantPack, sst08
  ],
  "Infant Cohort 3 Visit 5,EW Day 28": [
    infantPack, edta05, sst08, edta08
  ]
};

// =======================
// ALL EQUIPMENT (FOR NEW KIT/STUDY)
// =======================

const equipmentLibrary = [
  infantPack, edta05, edta08, sst05, sst08
];

// =======================
// SERVER
// =======================
app.get("/login", (req, res) => {
  res.send(`
    <h2>Login</h2>
    <form method="POST" action="/login">
      <input name="username" placeholder="Username" required /><br><br>
      <input name="password" type="password" placeholder="Password" required /><br><br>
      <button type="submit">Login</button>
    </form>
  `);
});
app.post("/login", (req, res) => {
  const { username, password } = req.body;

  if (username === "admin" && password === "1234") {
    req.session.loggedIn = true;
    res.redirect("/");
  } else {
    res.send("Invalid credentials");
  }
});

// LOGOUT
app.get("/logout", (req, res) => {
  req.session.destroy(() => {
    res.redirect("/login");
  });
});

// MAIN PAGE (PROTECTED)
app.get("/", (req, res) => {
  if (!req.session.loggedIn) {
    return res.redirect("/login");
  }

  res.send(`

<!DOCTYPE html>
<html>
<head>
<title>ICL Kit Laminate</title>

<style>

input {
  width: 100%;
  padding: 6px;
  margin-top: 4px;
  border: 1px solid #ccc;
  border-radius: 4px;
}

body {
  font-family: Arial;
  background: #f4f6f9;
  padding: 20px;
}

.container {
  max-width: 1200px;
  margin: auto;
  background: white;
  padding: 25px;
  border-radius: 10px;
}

.header {
  text-align: center;
}

.header img {
 width: 200px;
}

select {
  width: 100%;
  padding: 10px;
  margin-bottom: 10px;
}

table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 20px;
}

th {
  background: #2c3e50;
  color: white;
}

td, th {
  border: 1px solid #ccc;
  padding: 8px;
}

img {
  width: 70px;
}

textarea {
  width: 100%;
  height: 60px;
  border: none;
}

.approval {
  display: flex;
  justify-content: space-between;
  margin-top: 30px;
}

.box {
  width: 48%;
  border: 1px solid #ccc;
  padding: 10px;
}

</style>

</head>

<body>

<div class="container">

<button onclick="downloadExcel()">Download Excel</button>
<button onclick="window.location.href='/logout'">Logout</button>
<div class="header">
  <img src="/images/IC_Labs_Logo.png">
  <h1> </h1>
</div>

<label>Study</label>
<select id="studySelect" onchange="loadStudy(this)">
  <option value="">Select Study</option>
  <option value="ICL-CABNATe-FAMCRU">ICL-CABNATe-FAMCRU</option>
  <option value="NEW_STUDY">+ New Study</option>
</select>

<label>Kit</label>
<select id="kitSelect" onchange="loadKit(this)">
  <option value="">Select Kit</option>
  <option value="NEW_KIT">+ New Kit</option>
</select>

<table id="kitTable">
<tr>
<th>Equipment</th>
<th>Image</th>
<th>Qty</th>
<th>Description</th>
<th>Instructions</th>
</tr>
</table>

<div class="approval">

<div class="box">
<h3>Client Approval</h3>

<label>Approved By:</label>
<input type="text" id="clientApprovedBy"><br><br>

<label>Designation:</label>
<input type="text" id="clientDesignation"><br><br>

<label>Date:</label>
<input type="date" id="clientDate">
</div>

<div class="box">
<h3>IC Labs Review</h3>

<label>Checked By:</label>
<input type="text" id="checkedBy"><br><br>

<label>Approved By:</label>
<input type="text" id="iclApprovedBy"><br><br>

<label>Date:</label>
<input type="date" id="iclDate">
</div>
</div>

</div>

</div>



<script>

const studyToKitsMap = ${JSON.stringify(studyToKitsMap)};
const kitTemplates = ${JSON.stringify(kitTemplates)};
const equipmentLibrary = ${JSON.stringify(equipmentLibrary)};
let customStudyName = "";
let customKitName = "";

function loadStudy(select) {
  const kitSelect = document.getElementById("kitSelect");

  kitSelect.innerHTML = '<option value="">Select Kit</option><option value="NEW_KIT">+ New Kit</option>';
  clearTable();

 if (select.value === "NEW_STUDY") {
  const name = prompt("Enter new study name:");

  if (!name) {
    select.value = "";
    return;
  }

  customStudyName = name;

  const kitSelect = document.getElementById("kitSelect");
  kitSelect.innerHTML =
    '<option value="">Select Kit</option><option value="NEW_KIT">+ New Kit</option>';

  showEquipmentSelector();
  return;
}

  const kits = studyToKitsMap[select.value];
  if (!kits) return;

  kits.forEach(k => {
    kitSelect.innerHTML += '<option value="' + k + '">' + k + '</option>';
  });
}

function loadKit(select) {
  clearTable();

 if (select.value === "NEW_KIT") {
  const name = prompt("Enter new kit name:");

  if (!name) {
    select.value = "";
    return;
  }

  customKitName = name;

  showEquipmentSelector();
  return;
}

  const kit = kitTemplates[select.value];
  if (!kit) return;

  kit.forEach(addRow);
}

function showEquipmentSelector() {
  clearTable();

  const table = document.getElementById("kitTable");

  equipmentLibrary.forEach(item => {
    const row = document.createElement("tr");

    row.innerHTML =
      '<td><input type="checkbox" onchange="toggleQty(this)"> ' + item.name + '</td>' +
      '<td><img src="' + item.img + '"></td>' +
      '<td><input type="number" value="' + item.qty + '" disabled></td>' +
      '<td><textarea readonly>' + item.desc + '</textarea></td>' +
      '<td><textarea readonly>' + item.inst + '</textarea></td>';

    table.appendChild(row);
  });
}

function toggleQty(cb) {
  const row = cb.closest("tr");
  const input = row.querySelector("input[type=number]");
  input.disabled = !cb.checked;
}

function addRow(item) {
  const table = document.getElementById("kitTable");

  const row = document.createElement("tr");

  row.innerHTML =
    '<td>' + item.name + '</td>' +
    '<td><img src="' + item.img + '"></td>' +
    '<td><input type="number" value="' + item.qty + '"></td>' +
    '<td><textarea readonly>' + item.desc + '</textarea></td>' +
    '<td><textarea readonly>' + item.inst + '</textarea></td>';

  table.appendChild(row);
}

function clearTable() {
  document.getElementById("kitTable").innerHTML =
    '<tr><th>Equipment</th><th>Image</th><th>Qty</th><th>Description</th><th>Instructions</th></tr>';
}

function downloadExcel() {
  const kitItems = Array.from(document.querySelectorAll("#kitTable tr"))
  .slice(1)
  .filter(row => {
    const checkbox = row.querySelector("input[type=checkbox]");
    return checkbox ? checkbox.checked : true; 
  })
  .map(row => {
      const cells = row.children;

      const qtyInput = cells[2].querySelector("input");
      const qty = qtyInput ? qtyInput.value : cells[2].innerText;

      const imgEl = cells[1].querySelector("img");
      const img = imgEl ? imgEl.src : null;

      return {
  name: cells[0].innerText.trim(),
  img: img,
  qty: qty.trim(),
  desc: cells[3].querySelector("textarea")?.value.trim() || "",
  inst: cells[4].querySelector("textarea")?.value.trim() || "",
      };
    });

  fetch("/download-excel", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
   body: JSON.stringify({
studyText: customStudyName || document.getElementById("studySelect").value,
kitText: customKitName || document.getElementById("kitSelect").value,
  items: kitItems,

  // Approval fields
  clientApprovedBy: document.getElementById("clientApprovedBy").value,
  clientDesignation: document.getElementById("clientDesignation").value,
  clientDate: document.getElementById("clientDate").value,

  checkedBy: document.getElementById("checkedBy").value,
  iclApprovedBy: document.getElementById("iclApprovedBy").value,
  iclDate: document.getElementById("iclDate").value
}),

  })
  .then(res => res.blob())
  .then(blob => {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "kit-laminate.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
  })
  .catch(err => console.error("Error downloading Excel:", err));
}

</script>



</body>
</html>

`);

});





app.post("/download-excel", async (req, res) => {
  try {
    const {
  items,
  studyText,
  kitText,
  clientApprovedBy,
  clientDesignation,
  clientDate,
  checkedBy,
  iclApprovedBy,
  iclDate
} = req.body;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Kit Laminate");

    // ===== Add Study and Kit at the top
   
// ===== Study + Kit
const headerStudy = sheet.addRow([`Study: ${studyText || "N/A"}`]);
const headerKit = sheet.addRow([`Kit: ${kitText || "N/A"}`]);
sheet.addRow([]);
headerStudy.font = {bold: true};
headerKit.font = {bold: true};

// ===== Client Approval
const headerClientApproval = sheet.addRow(["Client Approval"]);
sheet.addRow([`Approved By: ${clientApprovedBy || ""}`]);
sheet.addRow([`Designation: ${clientDesignation || ""}`]);
sheet.addRow([`Date: ${clientDate || ""}`]);
sheet.addRow([]);
headerClientApproval.font = {bold: true};

// ===== IC Labs Review
const headerICLabsReviewer = sheet.addRow(["IC Labs Review"]);
sheet.addRow([`Checked By: ${checkedBy || ""}`]);
sheet.addRow([`Approved By: ${iclApprovedBy || ""}`]);
sheet.addRow([`Date: ${iclDate || ""}`]);
sheet.addRow([]);
headerICLabsReviewer.font = {bold: true};
	

    // ===== Add header row
const headerRow = sheet.addRow(["Equipment", "Image", "Qty", "Description", "Instructions"]);
headerRow.eachCell(cell => {
  cell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };
});

// ✅ Center headers
for (let j = 1; j <= 5; j++) {
  headerRow.getCell(j).alignment = {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true
  };
}

headerRow.font = { bold: true };
headerRow.height = 30;

    // ===== Add each item
// ===== Add each item
for (let i = 0; i < items.length; i++) {
  const item = items[i];

  // Add the row first with text
  const excelRow = sheet.addRow([
    item.name || "",
    "",               // Image column left blank, image is added separately
    item.qty || "",
    item.desc || "",
    item.inst || ""
  ]);
excelRow.eachCell(cell => {
  cell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };
});


// Center everything horizontally and vertically
for (let j = 1; j <= 5; j++) {
  excelRow.getCell(j).alignment = {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true
  };
}



  // Set row height to fit image + text
  excelRow.height = 80;

  // ===== Insert image if exists
  if (item.img) {
    try {
      const imageRes = await axios.get(item.img, { responseType: "arraybuffer" });
      const imageId = workbook.addImage({
        buffer: Buffer.from(imageRes.data, "binary"),
        extension: item.img.split(".").pop() // jpg or png
      });

      // Place image in column 2 only (ExcelJS is zero-indexed for coordinates)
sheet.addImage(imageId, {
  tl: { col: 1.2, row: excelRow.number - 0.8 }, // ✅ shifts image into center
  ext: { width: 100, height: 45 }
});
    } catch (err) {
      console.error("Error loading image:", item.img, err);
    }
  }
}

  sheet.columns = [
  { width: 30 }, // Equipment
  { width: 40 }, // Image (more space = better centering)
  { width: 10 }, // Qty
  { width: 45 }, // Description
  { width: 45 }  // Instructions
];



    // ===== Send file
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="kit-laminate.xlsx"');

    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).send("Error generating Excel");
  }
});

app.listen(PORT, () => {
  console.log("Running on http://localhost:" + PORT);
});