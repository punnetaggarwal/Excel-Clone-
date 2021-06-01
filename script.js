const ps = new PerfectScrollbar('#cells');

function findRowCol(ele) {
    let idArray = $(ele).attr("id").split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);

    return [rowId, colId];
}
function calColCode(n) {
    let str = "";

    while (n > 0) {
        let rem = n % 26;
        if (rem == 0) {
            str = 'Z' + str;
            n = Math.floor((n / 26)) - 1;
        } else {
            str = String.fromCharCode((rem - 1) + 65) + str;
            n = Math.floor((n / 26));
        }
    }
    return str;
}

for (let i = 1; i <= 100; i++) {
    let str = calColCode(i);
    $("#columns").append(`<div class="column-name">${str}</div>`);
    $("#rows").append(`<div class="row-name">${i}</div>`);
}

$("#cells").scroll(function () {
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
});

let saved = true;
let cellData = { "Sheet1": {} };
let selectedSheet = "Sheet1";
let totalSheets = 1;
let lastAddedSheetNumber = 1;
let mousemoved = false;
let startCellStored = false;
let startCell;
let endCell;
let defaultProperties = {
    "font-family": "Noto Sans",
    "font-size": "14",
    "bold": false,
    "italic": false,
    "underlined": false,
    "alignment": "left",
    "bgcolor": "#fff",
    "color": "#444",
    "text": "",
    "upStream": [],
    "downStream": [],
    "formula": ""
}

function loadNewSheet() {
    for (let i = 1; i <= 100; i++) {
        let row = $('<div class="cell-row"></div>');
        for (let j = 1; j <= 100; j++) {
            row.append(`<div id="row-${i}-col-${j}" class="input-cell"  contenteditable = "false"></div>`);
        }
        $("#cells").append(row);
    }
    addEventsToCells();
    addSheetTabEventListners();
}

loadNewSheet();

function addEventsToCells() {
    $(".input-cell").dblclick(function () {
        $(this).attr("contenteditable", "true");
        $(this).focus();
    });

    $(".input-cell").blur(function () {
        $(this).attr("contenteditable", "false");
        updatecellData("text", $(this).text());
    });


    $(".input-cell").click(function (e) {
        let [rowId, colId] = findRowCol(this);
        let [topCell, bottomCell, leftCell, rightCell] = findAdjacentCells(rowId, colId);
        if ($(this).hasClass("selected") && e.ctrlKey) {
            unselectCell(this, e, topCell, bottomCell, leftCell, rightCell, false);
        } else {
            selectCell(this, e, topCell, bottomCell, leftCell, rightCell, false);
        }
    });

    $(".input-cell").mousemove(function (event) {
        event.preventDefault();
        if (event.buttons == 1 && !event.ctrlKey) {
            $(".input-cell.selected").removeClass("selected topSelected bottomSelected leftSelected rightSelected");

            mousemoved = true;
            if (!startCellStored) {
                startCellStored = true;
                let [rowId, colId] = findRowCol(event.target);
                startCell = { rowId: rowId, colId: colId };
            } else {
                let [rowId, colId] = findRowCol(event.target);
                endCell = { rowId: rowId, colId: colId };
                selectAllBetweenTheRange(startCell, endCell);
            }
        } else {
            if (event.buttons == 0 && mousemoved) {
                startCellStored = false;
                mousemoved = false;
            }
        }
    });

}

function findAdjacentCells(rowId, colId) {
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);

    return [topCell, bottomCell, leftCell, rightCell];
}


function unselectCell(ele, e, topCell, bottomCell, leftCell, rightCell, mousemove) {
    if ((e.ctrlKey && $(ele).attr("contenteditable") == "false") || mousemove) {
        if ($(ele).hasClass("topSelected")) {
            topCell.removeClass("bottomSelected");
        }
        if ($(ele).hasClass("bottomSelected")) {
            bottomCell.removeClass("topSelected");
        }
        if ($(ele).hasClass("leftSelected")) {
            leftCell.removeClass("rightSelected");
        }
        if ($(ele).hasClass("rightSelected")) {
            rightCell.removeClass("leftSelected");
        }
        $(ele).removeClass("selected topSelected bottomSelected leftSelected rightSelected");
    }
}

function selectCell(ele, e, topCell, bottomCell, leftCell, rightCell, mousemove) {
    if (e.ctrlKey || mousemove) {
        let topSelected, bottomSelected, leftSelected, rightSelected;
        if (topCell) topSelected = topCell.hasClass("selected");
        if (bottomCell) bottomSelected = bottomCell.hasClass("selected");
        if (leftCell) leftSelected = leftCell.hasClass("selected");
        if (rightCell) rightSelected = rightCell.hasClass("selected");

        if (topSelected) {
            topCell.addClass("bottomSelected");
            $(ele).addClass("topSelected")
        }
        if (bottomSelected) {
            bottomCell.addClass("topSelected");
            $(ele).addClass("bottomSelected")
        }
        if (leftSelected) {
            leftCell.addClass("rightSelected");
            $(ele).addClass("leftSelected")
        }
        if (rightSelected) {
            rightCell.addClass("leftSelected");
            $(ele).addClass("rightSelected")
        }
    }
    else {
        $(".input-cell.selected").removeClass("selected topSelected bottomSelected leftSelected rightSelected");
    }
    $(ele).addClass("selected");
    changeHeader(findRowCol(ele));
}

function changeHeader([rowId, colId]) {
    let data;
    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
        data = cellData[selectedSheet][rowId - 1][colId - 1];
    } else {
        data = defaultProperties;
    }
    $("#font-family").val(data["font-family"]);
    $("#font-family").css("font-family", data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $(".alignment.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");
    addRemoveSelectFromFontStyle(data, "bold");
    addRemoveSelectFromFontStyle(data, "italic");
    addRemoveSelectFromFontStyle(data, "underlined");
    $("#fill-color-icon").css("border-bottom", `4px solid ${data.bgcolor}`);
    $("#text-color-icon").css("border-bottom", `4px solid ${data.color}`);
}

function addRemoveSelectFromFontStyle(data, property) {
    if (data[property]) {
        $(`#${property}`).addClass("selected");
    }
    else {
        $(`#${property}`).removeClass("selected");
    }
}

function selectAllBetweenTheRange(start, end) {
    for (let i = (start.rowId < end.rowId ? start.rowId : end.rowId); i <= (start.rowId < end.rowId ? end.rowId : start.rowId); i++) {
        for (let j = (start.colId < end.colId ? start.colId : end.colId); j <= (start.colId < end.colId ? end.colId : start.colId); j++) {
            let ele = $(`#row-${i}-col-${j}`)[0];                                              // to convert the jQuery thing into normal node
            let [topCell, bottomCell, leftCell, rightCell] = findAdjacentCells(i, j);  // we took the 0th index of the array 
            selectCell(ele, {}, topCell, bottomCell, leftCell, rightCell, true);
        }
    }
}

$(".menu-selector").change(function (e) {
    let value = $(this).val();
    let key = $(this).attr("id");
    if (key == "font-family") {
        $("#font-family").css(key, value);
    }
    if (!isNaN(value)) {
        value = parseInt(value);
    }
    $(".input-cell.selected").css(key, value);
    updatecellData(key, value);
});

$(".alignment").click(function (e) {
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    let alignment = $(this).attr("data-type");
    $(".input-cell.selected").css("text-align", alignment);
    updatecellData("alignment", alignment);
});


$("#bold").click(function (e) {
    setFontStyle(this, "bold", "font-weight", "bold");
});

$("#italic").click(function (e) {
    setFontStyle(this, "italic", "font-style", "italic");
});

$("#underlined").click(function (e) {
    setFontStyle(this, "underlined", "text-decoration", "underline");
});

function setFontStyle(ele, property, key, value) {
    if ($(ele).hasClass("selected")) {
        $(ele).removeClass("selected");
        $(".input-cell.selected").css(key, "");
        updatecellData(property, false);
    } else {
        $(ele).addClass("selected");
        $(".input-cell.selected").css(key, value);
        updatecellData(property, true);
    }
}

function updatecellData(property, value) {
    let prevCellData = JSON.stringify(cellData);
    if (value != defaultProperties[property]) {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = findRowCol(data);
            if (cellData[selectedSheet][rowId - 1] == undefined) {
                cellData[selectedSheet][rowId - 1] = {};
                cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties };
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
            } else {
                if (cellData[selectedSheet][rowId - 1][colId - 1] == undefined) {
                    cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties };
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                } else {
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                }
            }
        });
    } else {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = findRowCol(data);
            if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                if (JSON.stringify(cellData[selectedSheet][rowId - 1][colId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][rowId - 1][colId - 1];
                    if (Object.keys(cellData[selectedSheet][rowId - 1]).length == 0) {
                        delete cellData[selectedSheet][rowId - 1];
                    }
                }
            }
        });
    }
    if (saved && JSON.stringify(cellData) != prevCellData) {
        saved = false;
    }
}

$(".color-pick").colorPick({
    'initialColor': '#TYPECOLOR',
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function () {
        if (this.color != "#TYPECOLOR") {
            if (this.element.attr("id") == "fill-color") {
                $("#fill-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("background-color", this.color);
                updatecellData("bgcolor", this.color);
            }
            else {
                $("#text-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("color", this.color);
                updatecellData("color", this.color);
            }
        }
    }
});

$("#fill-color-icon, #text-color-icon").hover(function (e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 10);
});


$(".container").click(function (e) {
    $(".sheet-options-modal").remove();
});


$(".sheet-tab").click(function (e) {
    if (!$(this).hasClass("selected")) {
        selectSheet(this);
    }
});

function selectSheet(ele) {
    $(".sheet-tab.selected").removeClass("selected");
    $(ele).addClass("selected");
    emptySheet();
    selectedSheet = $(ele).text();
    loadSheet();
}

function emptySheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for (let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId])
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text("");
            cell.css({
                "font-family": "Noto Sans",
                "font-size": 14,
                "background-color": "#fff",
                "color": "#444",
                "font-weight": "",
                "font-style": "",
                "text-decoration": "",
                "text-align": "left"
            });
        }
    }
}

function loadSheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for (let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId])
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text(data[rowId][colId].text);
            cell.css({
                "font-family": data[rowId][colId]["font-family"],
                "font-size": data[rowId][colId]["font-size"] + "px",
                "background-color": data[rowId][colId]["bgcolor"],
                "color": data[rowId][colId].color,
                "font-weight": data[rowId][colId].bold ? "bold" : "",
                "font-style": data[rowId][colId].italic ? "italic" : "",
                "text-decoration": data[rowId][colId].underlined ? "underline" : "",
                "text-align": data[rowId][colId].alignment
            });
        }
    }
    $(`#row-1-col-1`).click();
}

$(".add-sheet").click(function (e) {
    emptySheet();
    totalSheets++;
    lastAddedSheetNumber++;
    cellData[`Sheet${lastAddedSheetNumber}`] = {};
    selectedSheet = `Sheet${lastAddedSheetNumber}`;
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheet-tab-container").append(
        `<div class="sheet-tab selected">Sheet${lastAddedSheetNumber}</div>`
    );
    $(".sheet-tab").click(function (e) {
        if (!$(this).hasClass("selected")) {
            selectSheet(this);
        }
    });
    addSheetTabEventListners();
    $(".sheet-tab.selected")[0].scrollIntoView();
    saved = false;
    $(`#row-1-col-1`).click();
});

function addSheetTabEventListners() {
    $(".sheet-tab.selected").bind("contextmenu", function (e) {
        e.preventDefault();
        $(".sheet-options-modal").remove();
        let modal = $(`<div class="sheet-options-modal">
                        <div class="option sheet-rename">Rename</div>
                        <div class="option sheet-delete">Delete</div>
                    </div>`);
        $(".container").append(modal);
        $(".sheet-options-modal").css({ "bottom": 0.04 * $(".container").height(), "left": e.pageX });
        $(".sheet-rename").click(function (e) {
            let renameModal = `<div class="sheet-rename-modal-parent">
            <div class="sheet-rename-modal">
                <div class="sheet-modal-title">
                    <span>Rename Sheet</span>
                </div>
                <div class="sheet-modal-input-container">
                    <span class="sheet-modal-input-title">Rename Sheet to:</span>
                    <input class="sheet-modal-input" type="text" >
                </div>
                <div class="sheet-modal-configration">
                    <div class="button ok-button">OK</div>
                    <div class="button cancel-button">Cancel</div>
                </div>
            </div>
            </div>`
            $(".container").append(renameModal);
            $(".cancel-button").click(function (e) {
                $(".sheet-rename-modal-parent").remove();
            });
            $(".ok-button").click(function (e) {
                renameSheet();
            });
            $(".sheet-modal-input").keypress(function (e) {
                if (e.key == "Enter") {
                    renameSheet();
                }
            });

        });

        $(".sheet-delete").click(function (e) {
            let deleteModal = `<div class="sheet-delete-modal-parent">
            <div class="sheet-delete-modal">
                <div class="sheet-delete-modal-title">
                    <span>Sheet Name</span>
                </div>
                <div class="sheet-delete-modal-container">
                    <span class="material-icons delete-symbol">delete</span>
                    <div class="warning-message">Are You Sure?</div>
                </div>
                <div class="sheet-delete-modal-configration">
                    <div class="button delete-button">Delete</div>
                    <div class="button cancel-button">Cancel</div>
                </div>
                    </div>
                </div>
            </div> `
            $(".container").append(deleteModal);
            $(".cancel-button").click(function (e) {
                $(".sheet-delete-modal-parent").remove();
            });
            $(".delete-button").click(function (e) {
                $(".sheet-delete-modal-parent").remove();
                if (totalSheets > 1) {
                    let keyArray = Object.keys(cellData);
                    let selectedSheetIndex = keyArray.indexOf(selectedSheet);
                    let currentSelectedSheet = $(".sheet-tab.selected");
                    if (selectedSheetIndex == 0) {
                        selectSheet(currentSelectedSheet.next()[0]);
                    } else {
                        selectSheet(currentSelectedSheet.prev()[0]);
                    }
                    delete cellData[currentSelectedSheet.text()];
                    currentSelectedSheet.remove();
                    totalSheets--;
                    saved = false;
                } else {

                }

            });
        });

        if (!$(this).hasClass("selected")) {
            selectSheet(this);
        }
    });

    $(".sheet-tab.selected").click(function (e) {
        if (!$(this).hasClass("selected")) {
            selectSheet(this);
        }
    });
}

function renameSheet() {
    let newSheetName = $(".sheet-modal-input").val();
    if (newSheetName && !Object.keys(cellData).includes(newSheetName)) {
        let newCellData = {};
        for (let i of Object.keys(cellData)) {
            if (i == selectedSheet) {
                newCellData[newSheetName] = cellData[i];
            } else {
                newCellData[i] = cellData[i];
            }
        }
        cellData = newCellData;
        selectedSheet = newSheetName;
        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-rename-modal-parent").remove();
    } else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(`
            <div class="error">Sheet name is not valid or Sheet already Exist!</div>
        `);
    }
    saved = false;
}

$(".left-scroller").click(function (e) {
    let keyArray = Object.keys(cellData);
    let currentSelectedSheetIndex = keyArray.indexOf(selectedSheet);
    let currentSelectedSheet = $(".sheet-tab.selected");
    if (currentSelectedSheetIndex > 0) {
        selectSheet(currentSelectedSheet.prev()[0]);
    }
});

$(".right-scroller").click(function (e) {
    let keyArray = Object.keys(cellData);
    let currentSelectedSheetIndex = keyArray.indexOf(selectedSheet);
    let currentSelectedSheet = $(".sheet-tab.selected");
    if (currentSelectedSheetIndex != keyArray.length - 1) {
        selectSheet(currentSelectedSheet.next()[0]);
    }
});

$("#menu-file").click(function (e) {
    let fileModal = $(`<div class="file-modal">
                            <div class="file-options-modal">
                                <div class="close">
                                    <div class="material-icons close-icon">arrow_circle_down</div>
                                    <div>Close</div>
                                </div>
                                <div class="new">
                                    <div class="material-icons new-icon">insert_drive_file</div>
                                    <div>New</div>
                                </div>
                                <div class="open">
                                    <div class="material-icons open-icon">folder_open</div>
                                    <div>Open</div>
                                </div>
                                <div class="save">
                                    <div class="material-icons save-icon">save</div>
                                    <div>Save</div>
                                </div>
                            </div>
                            <div class="file-recent-modal">
                            </div>
                            <div class="file-transparent-modal"></div>
                        </div>`);
    $(".container").append(fileModal);
    fileModal.animate({
        "width": "100vw"
    }, 300);
    $(".close,.file-transparent-modal,.new,.save, .open").click(function (e) {
        fileModal.animate({
            "width": "0vw"
        }, 300);
        setTimeout(() => {
            fileModal.remove();
        }, 299);
    });
    $(".new").click(function (e) {

        if (saved) {
            newFile();
        } else {
            
            $(".container").append(`<div class="sheet-delete-modal-parent">
            <div class="sheet-delete-modal">
                <div class="sheet-delete-modal-title">
                    <span>${$(".title-bar").text()}</span>
                </div>
                <div class="sheet-delete-modal-container">
                    <div class="warning-message">Do you want to save the changes?</div>
                </div>
                <div class="sheet-delete-modal-configration">
                    <div class="button delete-button">Yes</div>
                    <div class="button cancel-button">No</div>
                </div>
                    </div>
                </div>
            </div> `);
            $(".delete-button").click(function (e) {
                $(".sheet-delete-modal-parent").remove();
                saveFile(true);
            });
            $(".cancel-button").click(function (e) {
                $(".sheet-delete-modal-parent").remove();
                newFile();
                saved = true;
            })
        }
    });

    $(".save").click(function (e) {
        saveFile();
    })
    $(".open").click(function (e) {
        openFile();
    });
});

function newFile() {
    emptySheet();
    $(".sheet-tab").remove();
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet1</div>`);
    cellData = { "Sheet1": {} };
    selectedSheet = "Sheet1";
    totalSheets = 1;
    lastAddedSheetNumber = 1;
    addSheetTabEventListners();
    $("#row-1-col-1").click();
}

function saveFile(createNewFile) {
    if (!saved) {
        $(".container").append(`<div class="sheet-rename-modal-parent">
                                <div class="sheet-rename-modal">
                                    <div class="sheet-modal-title">
                                        <span>Save File</span>
                                    </div>
                                    <div class="sheet-modal-input-container">
                                        <span class="sheet-modal-input-title">File Name:</span>
                                        <input class="sheet-modal-input" value='${$(".title-bar").text()}' type="text" />
                                    </div>
                                    <div class="sheet-modal-configration">
                                        <div class="button ok-button">Save</div>
                                        <div class="button cancel-button">Cancel</div>
                                    </div>
                                </div>
                            </div>`);
        $(".ok-button").click(function (e) {
            let fileName = $(".sheet-modal-input").val();
            if (fileName) {
                let href = `data:application/json,${encodeURIComponent(JSON.stringify(cellData))}`;
                let a = $(`<a href=${href} download="${fileName}.json"></a>`);
                $(".container").append(a);
                a[0].click();
                a.remove();
                $(".sheet-rename-modal-parent").remove();
                saved = true;
                if (createNewFile) {
                    newFile();
                }
            }
        });
        $(".cancel-button").click(function (e) {
            $(".sheet-rename-modal-parent").remove();
            if (createNewFile) {
                newFile();
            }
        });
    }
}

function openFile() {
    let inputFile = $(`<input accept="application/json" type="file" />`);
    $(".container").append(inputFile);
    inputFile.click();
    inputFile.change(function (e) {
        let file = e.target.files[0];
        $(".title-bar").text(file.name.split(".json")[0]);
        let reader = new FileReader();
        reader.readAsText(file);
        reader.onload = function () {
            emptySheet();
            $(".sheet-tab").remove();
            cellData = JSON.parse(reader.result);
            let sheets = Object.keys(cellData);
            for (let i of sheets) {
                $(".sheet-tab-container").append(`<div class="sheet-tab selected">${i}</div>`)
            }
            addSheetTabEventListners();
            $(".sheet-tab").removeClass("selected");
            $($(".sheet-tab")[0]).addClass("selected");
            selectedSheet = sheets[0];
            totalSheets = sheets.length;
            lastAddedSheetNumber = totalSheets;
            loadSheet();
            inputFile.remove();
        }
    });
}


let clipBoard = { startCell: [], cellData: {} };
let contentCutted = false;

$("#cut, #copy").click(function (e) {
    if ($(this).text() == "content_cut") {
        contentCutted = true;
    }
    clipBoard.startCell = findRowCol($(".input-cell.selected")[0]);
    $(".input-cell.selected").each((index, data) => {
        let [rowId, colId] = findRowCol(data);
        if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
            if (!clipBoard.cellData[rowId]) {
                clipBoard.cellData[rowId] = {};
            }
            clipBoard.cellData[rowId][colId] = { ...cellData[selectedSheet][rowId - 1][colId - 1] };

        }
    });
});

$("#paste").click(function (e) {
    if (contentCutted) {
        emptySheet();
    }
    startCell = findRowCol($(".input-cell.selected")[0]);
    let rows = Object.keys(clipBoard.cellData);
    for (let i of rows) {
        let cols = Object.keys(clipBoard.cellData[i]);
        for (let j of cols) {
            if (contentCutted) {
                delete cellData[selectedSheet][i - 1][j - 1];
                if (Object.keys(cellData[selectedSheet][i - 1]).length == 0) {
                    delete cellData[selectedSheet][i - 1];
                }
            }
            let rowDistance = parseInt(i) - parseInt(clipBoard.startCell[0]);
            let colDistance = parseInt(j) - parseInt(clipBoard.startCell[1]);
            if (!cellData[selectedSheet][startCell[0] + rowDistance - 1]) {
                cellData[selectedSheet][startCell[0] + rowDistance - 1] = {};
            }
            cellData[selectedSheet][startCell[0] + rowDistance - 1][startCell[1] + colDistance - 1] = { ...clipBoard.cellData[i][j] };
        }
    }
    loadSheet();
    if (contentCutted) {
        contentCutted = false;
        clipBoard = { startCell: [], cellData: {} };
    }
});

$("#function-input").blur(function (e) {
    if ($(".input-cell.selected").length > 0) {
        let formula = $(this).text();
        if (formula == "") {
            alert("Please appy a formula");
        } else {
            let tempElements = formula.split(" ");
            let elements = [];
            for (let i of tempElements) {
                if (i.length > 1) {
                    i = i.replace("(", "");
                    i = i.replace(")", "");
                    if (!elements.includes(i)) {
                        elements.push(i);
                    }
                }
            }

            $(".input-cell.selected").each(function (index, data) {
                let [rowId, colId] = findRowCol(data);
                if (updateStreams(data, elements, false)) {
                    cellData[selectedSheet][rowId - 1][colId - 1].formula = formula;
                    let selfColCode = calColCode(colId);
                    evalFormula(selfColCode + rowId);
                } else {
                    alert("Formula is invalid!");
                }
            });
        }

    } else {
        alert("Please select a cell first to apply formula!");
    }
});

// function updateStreams(rowId, colId, elements, update, oldUpStreams) {
//     let selfColCode = calColCode(colId);

//     if (elements.includes(selfColCode + rowId)) {
//         return false;
//     }

//     if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
//         let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
//         let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;

//         for (let i of downStream) {
//             if (elements.includes(i)) {
//                 return false;
//             }
//         }

//         for (let i of downStream) {
//             let [calColId, calRowId] = codeToVal(i);
//             updateStreams(calRowId, calColId, elements, true, upStream);
//         }
//     }

//     if (!cellData[selectedSheet][rowId - 1]) {
//         cellData[selectedSheet][rowId - 1] = {};
//     }
//     if (!cellData[selectedSheet][rowId - 1][colId - 1]) {
//         cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upstream": [...elements], downStream: [] };
//     } else {
//         let upstream = { ...cellData[selectedSheet][rowId - 1][colId - 1].upstream };
//         if (update) {
//             for (let i of oldUpStreams) {
//                 let [calColId, calRowId] = codeToVal(i);
//                 let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
//                 cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);

//                 if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1] == JSON.stringify(defaultProperties))) {
//                     delete cellData[selectedSheet][calRowId - 1][calColId - 1];
//                     if (Object.keys(cellData[selectedSheet][calRowId - 1].length == 0)) {
//                         delete cellData[selectedSheet][calRowId - 1];
//                     }
//                 }
//                 index = cellData[selectedSheet][rowId - 1][colId - 1].upStream.indexOf(i);
//                 cellData[selectedSheet][rowId - 1][colId - 1].upStream.splice(index, 1);
//             }

//             for (let i of elements) {
//                 cellData[selectedSheet][rowId - 1][colId - 1].upStream.push(i);
//             }
//         } else {
//             for (let i of upstream) {
//                 let [calColId, calRowId] = codeToVal(i);
//                 let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
//                 cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);

//                 if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1] == JSON.stringify(defaultProperties))) {
//                     delete cellData[selectedSheet][calRowId - 1][calColId - 1];
//                     if (Object.keys(cellData[selectedSheet][calRowId - 1].length == 0)) {
//                         delete cellData[selectedSheet][calRowId - 1];
//                     }
//                 }
//             }
//             cellData[selectedSheet][rowId - 1][colId - 1].upstream = [...elements];
//         }

//     }

//     for (let i of elements) {
//         let [calRowId, calColId] = codeToVal(i);
//         if (!cellData[selectedSheet][calRowId - 1]) {
//             cellData[selectedSheet][calRowId - 1] = {};
//         }
//         if (!cellData[selectedSheet][calRowId - 1][calColId - 1]) {
//             cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upstream": [], downStream: [selfColCode + rowId] };
//         } else {
//             cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.push(selfColCode + rowId);
//         }
//     }

//     return true;
// }

// function evalFormula(code) {
//     let [rowId, colId] = codeToVal(code);
//     let formula = cellData[selectedSheet][rowId - 1][colId - 1].formula;
//     if (formula != "") {
//         let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
//         let upStreamValues = [];
//         for (let i of upStream) {
//             let [calRowId, calColId] = codeToVal(i);
//             let value;
//             if (cellData[selectedSheet][calRowId - 1][calColId - 1] == "") {
//                 value = 0;
//             } else {
//                 value = cellData[selectedSheet][calRowId - 1][calColId - 1].text;
//             }
//             upStreamValues.push(value);
//             formula.replace(upStream[i], upStreamValues[i]);
//         }
//         cellData[selectedSheet][rowId - 1][colId - 1].text = evel(formula);
//         loadSheet();
//     }

//     let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
//     for (let i = downStream.length - 1; i >= 0; i--) {
//         evalFormula(downStream[i]);
//     }
// }


function codeToValue (str) {
    let rowId, colId;
    let rowCode = "", colCode = "";
    for (let i = 0; i < str.length; i++) {
        let ch = str.charAt(i);
        if (isNaN(ch)) {
            colCode += ch;
        } else {
            rowCode += ch;
        }
    }
    let pow = 0;
    for (let i = colCode.length - 1; i >= 0; i--) {
        let charValue = str.charCodeAt(i) - 64;
        colId += Math.pow(26, pow) * charValue;
        pow++;
    }

    rowId = parseInt(rowCode);
    return [rowId, colId];
}


function updateStreams(ele, elements, update, oldUpstream) {
    let [rowId, colId] = findRowCol(ele);
    let selfColCode = calColCode(colId);
    if (elements.includes(selfColCode + rowId)) {
        return false;
    }
    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
        let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
        let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        for (let i of downStream) {
            if (elements.includes(i)) {
                return false;
            }
        }
        for (let i of downStream) {
            let [calRowId, calColId] = codeToValue(i);
            updateStreams($(`#row-${calRowId}-col-${calColId}`)[0], elements, true, upStream);
        }
    }

    if (!cellData[selectedSheet][rowId - 1]) {
        cellData[selectedSheet][rowId - 1] = {};
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    } else if (!cellData[selectedSheet][rowId - 1][colId - 1]) {
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    } else {

        let upStream = [...cellData[selectedSheet][rowId - 1][colId - 1].upStream];
        if (update) {
            for (let i of oldUpstream) {
                let [calRowId, calColId] = codeToValue(i);
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1]).length == 0) {
                        delete cellData[selectedSheet][calRowId - 1];
                    }
                }
                index = cellData[selectedSheet][rowId - 1][colId - 1].upStream.indexOf(i);
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.splice(index, 1);
            }
            for (let i of elements) {
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.push(i);
            }
        } else {
            for (let i of upStream) {
                let [calRowId, calColId] = codeToValue(i);
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1]).length == 0) {
                        delete cellData[selectedSheet][calRowId - 1];
                    }
                }
            }
            cellData[selectedSheet][rowId - 1][colId - 1].upStream = [...elements];
        }
    }

    for (let i of elements) {
        let [calRowId, calColId] = codeToValue(i);
        if (!cellData[selectedSheet][calRowId - 1]) {
            cellData[selectedSheet][calRowId - 1] = {};
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        } else if (!cellData[selectedSheet][calRowId - 1][calColId - 1]) {
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        } else {
            cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.push(selfColCode + rowId);
        }
    }
    
    return true;
}

function evalFormula(cell) {
    let [rowId, colId] = codeToValue(cell);
    let formula = cellData[selectedSheet][rowId - 1][colId - 1].formula;
    
    if (formula != "") {
        let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        let upStreamValue = [];
        for (let i in upStream) {
            let [calRowId, calColId] = codeToValue(upStream[i]);
            let value;
            if (cellData[selectedSheet][calRowId - 1][calColId - 1].text == "") {
                value = "0";
            }
             else {
                value = cellData[selectedSheet][calRowId - 1][calColId - 1].text;
            }
            upStreamValue.push(value);
            formula = formula.replace(upStream[i], upStreamValue[i]);
        }
        cellData[selectedSheet][rowId - 1][colId - 1].text = eval(formula);
        loadCurrentSheet();
    }
    let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
    for (let i = downStream.length - 1; i >= 0; i--) {
        evalFormula(downStream[i]);
    }

}