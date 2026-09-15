// Javascript for FLS
// Last updated: 12/09/26

let backgroundAdded = false;
let personAdded = false;

// Get input value
function getValue(id) {
const element = document.getElementById(id);
if (!element) {
return "";
}
return element.value.trim();
}

// Escape HTML
function escapeHTML(value) {
const div = document.createElement("div");
div.textContent = value;
return div.innerHTML;

}

// Status Message
function showStatus(message) {
document.getElementById("status").textContent = message;
}

// Add background information
function addBackground() {
    if (backgroundAdded) {
    showStatus("Background information has already been added.");
    return;
    }

    const values = [
        getValue("organiser"),
        getValue("jobSite"),
        getValue("surburb"),
        getValue("builder"),
        getValue("subContractor"),
        getValue("trade")
    ];

    if (values.every(value => value === "")) {
        alert("Please enter the background information first.");
        return;
    }

    const row = document.getElementById("backgroundInputRow");

    row.innerHTML = `
        <td class="locked-cell">
            ${escapeHTML(values[0])}
        </td>

        <td class="locked-cell">
            ${escapeHTML(values[1])}
        </td>

        <td class="locked-cell">
            ${escapeHTML(values[2])}
        </td>

        <td class="locked-cell">
            ${escapeHTML(values[3])}
        </td>

        <td class="locked-cell">
            ${escapeHTML(values[4])}
        </td>

        <td class="locked-cell">
            ${escapeHTML(values[5])}
        </td>

        <td class="locked-cell">
            ✓
        </td>
    `;

    backgroundAdded = true;

    const button = document.getElementById("backgroundButton");

    button.textContent = "Background info added ✓";
    button.disabled = true;
    button.style.opacity = "0.8";
    button.style.cursor = "default";

    showStatus("Background information added. You can now add member details.");

}

// Create/add member input row

function createPersonInputRow() {
    const tbody = document.getElementById("peopleBody");
    const row = document.createElement("tr");
    row.className = "person-input-row";
    row.innerHTML = `
    <td>
        <input class="input surname"
               type="text"
               placeholder="Surname">
    </td>

    <td>
        <input class="input firstName"
               type="text"
               placeholder="First Name">
    </td>

    <td>
        <input class="input dateOfBirth"
               type="text"
               placeholder="Date of Birth">
    </td>

    <td>
        <input class="input unionNo"
               type="text"
               placeholder="Union No.">
    </td>

    <td>
        <input class="input financialStatus"
               type="text"
               placeholder="Financial Status">
    </td>

    <td>
        <input class="input notes"
               type="text"
               placeholder="Notes">
    </td>

    <td class="remove-cell">
        <button
            class="remove-button"
            type="button"
            onclick="this.closest('tr').remove()">
            ×
        </button>
    </td>
`;
tbody.appendChild(row);
return row;
}


// Add person
function addPerson() {
    if (!backgroundAdded) {
        alert( "Please click 'Add background data' first.");
        return;
    }

    const tbody = document.getElementById("peopleBody");

    let inputRow = tbody.querySelector(".person-input-row");
    if (!inputRow) {
        inputRow = createPersonInputRow();
    }

    personAdded = true;

    const surname = inputRow.querySelector(".surname").value.trim();

    const firstName = inputRow.querySelector(".firstName").value.trim();

    const dateOfBirth = inputRow.querySelector(".dateOfBirth").value.trim();

    const unionNo = inputRow.querySelector(".unionNo").value.trim();

    const financialStatus =
        inputRow.querySelector(".financialStatus").value.trim();

    const notes = inputRow.querySelector(".notes").value.trim();

    if (
        surname === "" ||
        firstName === "" ||
        dateOfBirth === ""
    ) {

        alert( "Please enter Surname, First Name and Date of Birth.");
        return;
    }


//   Convert input fields into normal text.


inputRow.innerHTML = `
    <td>
        ${escapeHTML(surname)}
    </td>

    <td>
        ${escapeHTML(firstName)}
    </td>

    <td>
        ${escapeHTML(dateOfBirth)}
    </td>

    <td>
        ${escapeHTML(unionNo)}
    </td>

    <td>
        ${escapeHTML(financialStatus)}
    </td>

    <td>
        ${escapeHTML(notes)}
    </td>

    <td class="remove-cell">

        <button
            class="remove-button"
            type="button"
            onclick="this.closest('tr').remove()">
            ×
        </button>

    </td>
`;

inputRow.className = "person-text-row";


// Automatically create another input row.


const newRow =
    createPersonInputRow();

newRow
    .querySelector(".surname")
    .focus();

showStatus(
    surname +
    " " +
    firstName +
    " added to the list."
);

}

// Delete last person
function deleteLastPerson() {

const tbody = document.getElementById("peopleBody");

const rows =
    tbody.querySelectorAll("tr");

if (rows.length === 0) {

    showStatus(
        "There are no people to delete."
    );

    return;
}

rows[rows.length - 1].remove();

/*
   Always leave an input row
   available for the next person.
*/

if (!tbody.querySelector(".person-input-row")) {
    createPersonInputRow();
}

showStatus("Last row removed.");

}

// EXPORT EXCEL SHEET

function exportToExcel() {


 //  Check that XLSX actually loaded.


if (typeof XLSX === "undefined") {

    alert(
        "The Excel library did not load. " +
        "Please make sure you are connected to the internet " +
        "and reload the page."
    );

    return;
}

if (!personAdded) {
    alert(
        "Please add at least ONE member"
    );
    return;
}


//   Background info must be converted first.
if (!backgroundAdded) {
    alert(
        "Please click 'Add background data' before exporting."
    );
    return;
}

// Background Data

const backgroundRow =
    document.getElementById("backgroundInputRow");

const backgroundCells =
    backgroundRow.querySelectorAll("td");

const backgroundData = [

    backgroundCells[0].textContent.trim(),
    backgroundCells[1].textContent.trim(),
    backgroundCells[2].textContent.trim(),
    backgroundCells[3].textContent.trim(),
    backgroundCells[4].textContent.trim(),
    backgroundCells[5].textContent.trim()

];

/* ========================================================
   EXCEL DATA
   ======================================================== */

const data = [
    [
        "FINANCIAL STANDING LIST",
        "",
        "",
        "",
        "",
        ""
    ],

    [
        "ORGANISER/DELEGATE",
        "JOB SITE",
        "SURBURB",
        "BUILDER",
        "SUB-CONTRACTOR",
        "TRADE"
    ],

    backgroundData,
    [
        "SURNAME",
        "FIRST NAME",
        "DATE OF BIRTH",
        "UNION NO.",
        "FINANCIAL STATUS ($)",
        "NOTES FROM MEMBERSHIP FILE"
    ]
];

/* ========================================================
   PEOPLE
   ======================================================== */

const personRows = document.querySelectorAll( "#peopleBody .person-text-row");

personRows.forEach(function(row) {

    const cells =
        row.querySelectorAll("td");

    if (cells.length < 6) {
        return;
    }

    data.push([

        cells[0].textContent.trim(),
        cells[1].textContent.trim(),
        cells[2].textContent.trim(),
        cells[3].textContent.trim(),
        cells[4].textContent.trim(),
        cells[5].textContent.trim()

    ]);

});

// Create Worksheet

const worksheet =
    XLSX.utils.aoa_to_sheet(data);

const workbook =
    XLSX.utils.book_new();

// Merge title

worksheet["!merges"] = [
    {
        s: {
            r: 0,
            c: 0
        },

        e: {
            r: 0,
            c: 5
        }
    }
];

// Excel collumn widths

worksheet["!cols"] = [
    { wch: 22 },
    { wch: 20 },
    { wch: 18 },
    { wch: 18 },
    { wch: 23 },
    { wch: 40 }
];

// Excel styling

const titleStyle = {

    font: {
        name: "Arial",
        sz: 16,
        bold: true,
        color: {
        rgb: "FFFFFF"
        }
    },

    fill: {
        fgColor: {
            rgb: "343A40"
        }
    },

    alignment: {
        horizontal: "center",
        vertical: "center"
    }

};



const headingStyle = {

    font: {
        name: "Arial",
        sz: 12,
        bold: true,
        color: {
            rgb: "343A40"
        }
    },

    fill: {
        fgColor: {
            rgb: "E9ECEF"
        }
    },

    alignment: {
        horizontal: "center",
        vertical: "center",
        wrapText: true
    }

};

const backgroundStyle = {

    font: {
        name: "Arial",
        sz: 10,
        bold: false,
        color: {
            rgb: "343A40"
        }
    },

    fill: {
        fgColor: {
            rgb: "F8F9FA"
        }
    },

    alignment: {
        horizontal: "center",
        vertical: "center",
        wrapText: true
    }

};

const dataStyle = {

    font: {
        name: "Arial",
        sz: 10,
        bold: false,
        color: {
            rgb: "343A40"
        }
    },

    alignment: {
        vertical: "center",
        wrapText: true
    },

    border: {

        bottom: {
            style: "thin",
            color: {
                rgb: "E1E4E8"
            }
        }

    }

};

// Apply excel style

const range =
    XLSX.utils.decode_range(
        worksheet["!ref"]
    );

for (
    let R = range.s.r;
    R <= range.e.r;
    R++
) {

    for (
        let C = range.s.c;
        C <= range.e.c;
        C++
    ) {

        const address =
            XLSX.utils.encode_cell({
                r: R,
                c: C
            });

        const cell =
            worksheet[address];

        if (!cell) {
            continue;
        }

        /*
           Row 0 = title
        */

        if (R === 0) {

            cell.s =
                titleStyle;

        }

        /*
           Row 1 = background headings
           Row 3 = member headings

           BOTH are bold.
        */

        else if (
            R === 1 ||
            R === 3
        ) {

            cell.s =
                headingStyle;

        }

        /*
           Row 2 = background information
        */

        else if (R === 2) {

            cell.s =
                backgroundStyle;

        }

        /*
           Row 4 onwards = people
        */

        else if (R >= 4) {

            cell.s =
                dataStyle;

        }

    }

}

// ROw heights

worksheet["!rows"] = [

    { hpt: 30 },
    { hpt: 28 },
    { hpt: 26 },
    { hpt: 34 }

];

// Add worksgeet

XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "FSL Report"
);

// Download

try {

    XLSX.writeFile(
        workbook,
        "Financial_Standing_List.xlsx"
    );

    showStatus(
        "Excel file exported successfully."
    );

}

catch (error) {

    console.error(error);
    alert(
        "The Excel file could not be downloaded. " + "Please check your browser's download settings."
    );

}

}

// Start with one empty person row
createPersonInputRow();
