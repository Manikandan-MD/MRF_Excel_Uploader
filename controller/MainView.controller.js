sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox"
], (Controller, JSONModel, MessageToast, MessageBox) => {
    "use strict";

    return Controller.extend("mrpexceluploader.controller.MainView", {

        onInit: function () {
            let oJsonModel = new JSONModel({ results: [] });
            this.getView().setModel(oJsonModel, "excelModel");
            this._iUniqueSeq = 0;

            if (typeof XLSX === "undefined") {
                let oScript = document.createElement("script");
                oScript.src = "https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js";
                oScript.onload = function () {
                    console.info("SheetJS loaded dynamically.");
                };
                oScript.onerror = function () {
                    MessageBox.error(
                        "Failed to load SheetJS library.\n\nCheck your internet connection, or add this to index.html:\n\n" +
                        '<script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"><\/script>',
                        { title: "Library Load Error" }
                    );
                };
                document.head.appendChild(oScript);
            }
        },

        _generateUniqueId: function () {
            this._iUniqueSeq += 1;
            return String(this._iUniqueSeq).padStart(2, "0");
        },

        onSendToTable: function () {
            if (typeof XLSX === "undefined") {
                MessageToast.show("SheetJS (XLSX) library is not loaded. Check index.html.");
                return;
            }

            let oUploader = this.byId("fileUploader");
            let oFile = oUploader.oFileUpload.files[0];

            if (!oFile) {
                MessageToast.show("Please choose an Excel file first.");
                return;
            }

            let sName = oFile.name.toLowerCase();
            if (!sName.endsWith(".xlsx") && !sName.endsWith(".xls")) {
                MessageToast.show("Only .xlsx / .xls files are supported.");
                return;
            }

            let oReader = new FileReader();

            oReader.onload = function (e) {
                try {
                    let sData = e.target.result;
                    let workbook = XLSX.read(sData, { type: "binary" });
                    let sSheetName = workbook.SheetNames[0];
                    let sheet = workbook.Sheets[sSheetName];
                    let excelData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

                    if (!excelData || excelData.length === 0) {
                        MessageToast.show("No data found in the Excel sheet.");
                        return;
                    }

                    this._iUniqueSeq = 0;

                    let aResults = excelData.map(function (item) {
                        return {
                            UNIQUE_ID:       this._generateUniqueId(),
                            MATERIAL:        item["Material"]        || "",
                            PLANT:           item["Plant"]           || "",
                            STOCK_CATEGORY:  item["Stock Category"]  || "",
                            WBS_ORDER:       item["WBS / Order"]     || "",
                            SHEET_MATERIAL:  item["Sheet Material"]  || "",
                            METAL_GRADE:     item["Metal Grade"]     || "",
                            SHEET_THICKNESS: item["Sheet Thickness"] || "",
                            RM_LENGTH:       item["RM Length"]       || "",
                            RM_WIDTH:        item["RM Width"]        || "",
                            TOTAL_RM_QTY:    item["Total RM Qty"]    || "",
                            A:               item["A"]               || "",
                            B:               item["B"]               || "",
                            C:               item["C"]               || "",
                            D:               item["D"]               || ""
                        };
                    }.bind(this));

                    this.getView().getModel("excelModel").setProperty("/results", aResults);

                    let oRowCount = this.byId("idRowCountText");
                    if (oRowCount) {
                        oRowCount.setText("Total Rows : " + aResults.length);
                    }

                    MessageToast.show(aResults.length + " row(s) loaded into table.");

                } catch (err) {
                    MessageBox.error("Failed to read Excel file:\n" + err.message, { title: "Read Error" });
                }
            }.bind(this);

            oReader.onerror = function () {
                MessageBox.error("Could not read the file. Please try again.", { title: "File Error" });
            };

            oReader.readAsBinaryString(oFile);
        },

        onDownloadTemplate: function () {
            if (typeof XLSX === "undefined") {
                MessageBox.error("SheetJS library is not loaded.", { title: "Library Missing" });
                return;
            }

            let aHeaders = [
                "Material", "Plant", "Stock Category", "WBS / Order",
                "Sheet Material", "Metal Grade", "Sheet Thickness",
                "RM Length", "RM Width", "Total RM Qty", "A", "B", "C", "D"
            ];

            let aResults = this.getView().getModel("excelModel").getProperty("/results");
            let aExcelRows = [];

            if (aResults && aResults.length > 0) {
                aResults.slice(0, 2).forEach(function (oRow) {
                    aExcelRows.push({
                        "Material":        oRow.MATERIAL        || "",
                        "Plant":           oRow.PLANT           || "",
                        "Stock Category":  oRow.STOCK_CATEGORY  || "",
                        "WBS / Order":     oRow.WBS_ORDER       || "",
                        "Sheet Material":  oRow.SHEET_MATERIAL  || "",
                        "Metal Grade":     oRow.METAL_GRADE     || "",
                        "Sheet Thickness": oRow.SHEET_THICKNESS || "",
                        "RM Length":       oRow.RM_LENGTH       || "",
                        "RM Width":        oRow.RM_WIDTH        || "",
                        "Total RM Qty":    oRow.TOTAL_RM_QTY    || "",
                        "A":               oRow.A               || "",
                        "B":               oRow.B               || "",
                        "C":               oRow.C               || "",
                        "D":               oRow.D               || ""
                    });
                });
                MessageToast.show("Template downloaded with sample rows.");
            } else {
                let oEmptyRow = {};
                aHeaders.forEach(function (h) { oEmptyRow[h] = ""; });
                aExcelRows.push(oEmptyRow);
                MessageToast.show("No data loaded. Downloaded headers-only template.");
            }

            let oWorksheet = XLSX.utils.json_to_sheet(aExcelRows, { header: aHeaders });
            oWorksheet["!cols"] = [
                { wch: 20 }, { wch: 8  }, { wch: 28 }, { wch: 18 },
                { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 12 },
                { wch: 12 }, { wch: 14 }, { wch: 6  }, { wch: 6  },
                { wch: 6  }, { wch: 6  }
            ];

            let oWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(oWorkbook, oWorksheet, "MRP Upload Template");
            XLSX.writeFile(oWorkbook, "MRP_Excel_Upload_Template.xlsx");
        },

        onProcess: function () {
            let aAllData = this.getView().getModel("excelModel").getProperty("/results");

            if (!aAllData || aAllData.length === 0) {
                MessageBox.error(
                    "No line items found.\nPlease upload an Excel file with data.",
                    { title: "No Data" }
                );
                return;
            }

            let aValidRows = aAllData.filter(function (oRow) {
                return oRow.MATERIAL        || oRow.PLANT          ||
                       oRow.STOCK_CATEGORY  || oRow.WBS_ORDER      ||
                       oRow.SHEET_MATERIAL  || oRow.METAL_GRADE    ||
                       oRow.SHEET_THICKNESS || oRow.RM_LENGTH      ||
                       oRow.RM_WIDTH        || oRow.TOTAL_RM_QTY;
            });

            if (aValidRows.length === 0) {
                MessageBox.error(
                    "Excel file contains column headers only.\nPlease add data rows and upload again.",
                    { title: "Empty Data" }
                );
                return;
            }

            MessageBox.confirm(
                "Process all " + aValidRows.length + " line item(s) to backend?",
                {
                    title: "Confirm Process",
                    actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                    emphasizedAction: MessageBox.Action.YES,
                    onClose: function (sAction) {
                        if (sAction === MessageBox.Action.YES) {
                            this._callBackend(aValidRows);
                        }
                    }.bind(this)
                }
            );
        },

        _callBackend: function (aAllData) {

            let oModel = this.getOwnerComponent().getModel();

            if (!oModel) {
                MessageBox.error("OData model not found.", { title: "Model Error" });
                return;
            }

            jQuery.ajax({
                url: "/sap/opu/odata/sap/ZCB_SB_METAL/ZCB_I_METAL",
                method: "GET",
                headers: {
                    "X-CSRF-Token": "Fetch",
                    "Accept": "application/json"
                },
                success: function (oData, sStatus, oXHR) {
                    let sCsrfToken = oXHR.getResponseHeader("X-CSRF-Token");
                    if (!sCsrfToken) {
                        MessageBox.error("CSRF token not returned by server.", { title: "Auth Error" });
                        return;
                    }
                    oModel.oHeaders["X-CSRF-Token"] = sCsrfToken;
                    this._submitBatch(oModel, aAllData);
                }.bind(this),
                error: function (oXHR) {
                    MessageBox.error(
                        "CSRF token fetch failed: " + oXHR.status + " " + oXHR.statusText,
                        { title: "Auth Error" }
                    );
                }
            });
        },

        // _sanitize: function (val) {
        //     if (typeof val !== "string") return val;
        //     // Remove non-printable / hidden characters from Excel cells
        //     return val.replace(/[\x00-\x1F\x7F]/g, "").trim();
        // },

        // _submitBatch: function (oModel, aAllData) {

        //     let iSuccess = 0;
        //     let iError   = 0;
        //     let iTotal   = aAllData.length;

        //     aAllData.forEach(function (oRow,ind) {
        //         // if (ind !== 0) return;

        //         let oPayload = {
        //             SlNoQr:         this._sanitize(oRow.UNIQUE_ID        || ""),
        //             Material:       this._sanitize(oRow.MATERIAL         || ""),
        //             Plant:          this._sanitize(oRow.PLANT            || ""),
        //             StockCategory:  this._sanitize(oRow.STOCK_CATEGORY   || ""),
        //             WbsOrder:       this._sanitize(oRow.WBS_ORDER        || ""),
        //             SheetMaterial:  this._sanitize(oRow.SHEET_MATERIAL   || ""),
        //             MetalGrade:     this._sanitize(oRow.METAL_GRADE      || ""),
        //             SheetThickness: parseFloat(oRow.SHEET_THICKNESS) || 0,
        //             RmLength:       parseFloat(oRow.RM_LENGTH)       || 0,
        //             RmWidth:        parseFloat(oRow.RM_WIDTH)        || 0,
        //             TotalRmQty:     parseFloat(oRow.TOTAL_RM_QTY)    || 0,
        //             CutA:           parseFloat(oRow.A)               || 0,
        //             CutB:           parseFloat(oRow.B)               || 0,
        //             CutC:           parseFloat(oRow.C)               || 0,
        //             CutD:           parseFloat(oRow.D)               || 0
        //         };

        //         oModel.create("/ZCB_I_METAL", oPayload, {
        //             success: function (oData, oResponse) {
        //                 iSuccess += 1;
        //                 if (iSuccess + iError === iTotal) {
        //                     MessageToast.show(iSuccess + " of " + iTotal + " record(s) posted successfully!");
        //                 }
        //             },
        //             error: function (oError) {
        //                 iError += 1;
        //                 let sMsg = "";
        //                 try {
        //                     let oResp = JSON.parse(oError.responseText);
        //                     sMsg = oResp.error?.message?.value || oError.responseText;
        //                 } catch (e) {
        //                     sMsg = oError.responseText || oError.message || "Unknown error";
        //                 }
        //                 MessageBox.error("Failed to post row:\n" + sMsg, { title: "Backend Error" });
        //             }
        //         });

        //     }.bind(this));
        // }

        _submitBatch: function (oModel, aAllData) {

            let iSuccess = 0;
            let iError   = 0;
            let iTotal   = aAllData.length;

            // Force any value to a clean string
            const str = (val) => String(val === null || val === undefined ? "" : val)
                                    .replace(/[\x00-\x1F\x7F]/g, "")
                                    .trim();

            aAllData.forEach(function (oRow, ind) {

                let oPayload = {
                    SlNoQr:         str(oRow.UNIQUE_ID),
                    Material:       str(oRow.MATERIAL),
                    Plant:          str(oRow.PLANT),
                    StockCategory:  str(oRow.STOCK_CATEGORY),
                    WbsOrder:       str(oRow.WBS_ORDER),
                    SheetMaterial:  str(oRow.SHEET_MATERIAL),
                    MetalGrade:     str(oRow.METAL_GRADE),
                    SheetThickness: str(oRow.SHEET_THICKNESS) || 0,
                    RmLength:       str(oRow.RM_LENGTH)       || 0,
                    RmWidth:        str(oRow.RM_WIDTH)        || 0,
                    TotalRmQty:     str(oRow.TOTAL_RM_QTY)   || 0,
                    // CutA:           str(oRow.A)               || 0,
                    // CutB:           str(oRow.B)               || 0,
                    // CutC:           str(oRow.C)               || 0,
                    // CutD:           str(oRow.D)               || 0
                };

                oModel.create("/ZCB_I_METAL", oPayload, {
                    success: function (oData, oResponse) {
                        iSuccess += 1;
                        if (iSuccess + iError === iTotal) {
                            MessageToast.show(iSuccess + " of " + iTotal + " record(s) posted successfully!");
                        }
                    },
                    error: function (oError) {
                        iError += 1;
                        let sMsg = "";
                        try {
                            let oResp = JSON.parse(oError.responseText);
                            sMsg = oResp.error?.message?.value || oError.responseText;
                        } catch (e) {
                            sMsg = oError.responseText || oError.message || "Unknown error";
                        }
                        MessageBox.error("Failed to post row:\n" + sMsg, { title: "Backend Error" });
                    }
                });

            }.bind(this));
        },

    });
});