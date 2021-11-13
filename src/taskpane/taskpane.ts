/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("showAll").onclick = showAll;
    document.getElementById("showPrecedents").onclick = showPrecedents;
    document.getElementById("showDependents").onclick = showDependents;
    document.getElementById("clear").onclick = clear;
  }
});

export async function clear() {
  try {
    await Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      const ranges = activeSheet.getRanges();
      ranges.load("format/fill/color");
      return context.sync().then(function () {
        // Check whether affected by other functions
        ranges.format.fill.color = "White";
      });
    });
  } catch (error) {
    console.error(error);
  }
}

const colorPicker = () => {
  const letters = "BCDEF".split("");
  let color = "";
  for (let i = 0; i < 6; i++) {
    color += letters[Math.floor(Math.random() * letters.length)];
  }
  return color;
};

const message = (id: any, message: any) => {
  return (document.getElementById(id).innerText = message);
};

export async function showAll() {
  try {
    await Excel.run(async (context) => {
      const targetCell = context.workbook.getActiveCell();

      const directPrec = targetCell.getDirectPrecedents();
      const directDep = targetCell.getDirectDependents();
      targetCell.load("address");
      directPrec.areas.load("address");
      directDep.areas.load("address");
      const color = colorPicker();
      return context.sync().then(function () {
        for (let i = 0; i < directPrec.areas.items.length; i++) {
          directPrec.areas.items[i].format.fill.color = color;
        }
        for (let i = 0; i < directDep.areas.items.length; i++) {
          const currentArea = directDep.areas.items[i];
          currentArea.format.fill.color = color;
          currentArea.format.fill.tintAndShade = 0.7;
        }
      });
    });
  } catch (error) {
    if (error.code === "ItemNotFound") {
      message("message", "Relevant cells not found");
    }
    console.error(error);
  }
}
export async function showPrecedents() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const targetCell = context.workbook.getActiveCell();

      const directPrec = targetCell.getDirectPrecedents();
      targetCell.load("address");
      directPrec.areas.load("address");
      const previousAddress = directPrec;
      const color = colorPicker();
      return context.sync().then(function () {
        for (let i = 0; i < directPrec.areas.items.length; i++) {
          directPrec.areas.items[i].format.fill.color = color;
        }
        sheet.onSelectionChanged.add(async (event) => {
          previousAddress.load("address");
          return context.sync().then(function () {
            for (let i = 0; i < previousAddress.areas.items.length; i++) {
              event.type ?? (previousAddress.areas.items[i].format.fill.color = "white");
            }
          });
        });
      });
    });
  } catch (error) {
    if (error.code === "ItemNotFound") {
      message("message", "Precedent cells not found");
    }
    console.error(error);
  }
}

export async function showDependents() {
  try {
    await Excel.run(async (context) => {
      const targetCell = context.workbook.getActiveCell();

      const directDep = targetCell.getDirectDependents();
      targetCell.load("address");
      directDep.areas.load("address");
      const color = colorPicker();
      return context.sync().then(function () {
        for (let i = 0; i < directDep.areas.items.length; i++) {
          const currentArea = directDep.areas.items[i];
          currentArea.format.fill.color = color;
          currentArea.format.fill.tintAndShade = 0.7;
        }
      });
    });
  } catch (error) {
    if (error.code === "ItemNotFound") {
      message("message", "Dependent cells not found");
    }
    console.error(error);
  }
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRanges();
      const formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
      context.sync().then(function () {
        if (formulaRanges.isNullObject) {
          message("message", "No formulas found");
          return;
        }
        formulaRanges.format.fill.color = "#f5a8bb";
        context.sync();
      });
      return context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
