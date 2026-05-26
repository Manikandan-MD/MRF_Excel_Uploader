sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator"
], (Controller, JSONModel, MessageToast, MessageBox, Filter, FilterOperator) => {
    "use strict";

    return Controller.extend("mrpexceluploader.controller.MainView", {

        onInit: function () {
            let oJsonModel = new JSONModel({ results: [] });
            this.getView().setModel(oJsonModel, "excelModel");
            this.getView().setModel(new JSONModel({ count: 0 }), "reportCount");
            this.getView().setModel(new JSONModel({ results: [] }), "pipeExcelModel");
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

            this._showContainer("idUploadContainer");
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
                        MessageBox.error(
                            "No data found in the Excel sheet.\nPlease add data rows and upload again.",
                            { title: "Empty Sheet" }
                        );
                        return;
                    }

                    let aNonEmptyRows = excelData.filter(function (item) {
                        return Object.values(item).some(function (val) {
                            return val !== null && val !== undefined && String(val).trim() !== "";
                        });
                    });

                    if (aNonEmptyRows.length === 0) {
                        MessageBox.error(
                            "Excel file contains column headers only.\nPlease add data rows and upload again.",
                            { title: "No Data Rows" }
                        );
                        return;
                    }

                    this._iUniqueSeq = 0;

                    let aResults = excelData.map(function (item) {
                        return {
                            UNIQUE_ID: this._generateUniqueId(),
                            MATERIAL: item["Material"] || "",
                            PLANT: item["Plant"] || "",
                            STOCK_CATEGORY: item["Stock Category"] || "",
                            WBS_ORDER: item["WBS / Order"] || "",
                            SHEET_MATERIAL: item["Sheet Material"] || "",
                            METAL_GRADE: item["Metal Grade"] || "",
                            SHEET_THICKNESS: item["Sheet Thickness"] || "",
                            RM_LENGTH: item["RM Length"] || "",
                            RM_WIDTH: item["RM Width"] || "",
                            TOTAL_RM_QTY: item["Total RM Qty"] || "",
                            A: item["A"] || "",
                            B: item["B"] || "",
                            C: item["C"] || "",
                            D: item["D"] || ""
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
                        "Material": oRow.MATERIAL || "",
                        "Plant": oRow.PLANT || "",
                        "Stock Category": oRow.STOCK_CATEGORY || "",
                        "WBS / Order": oRow.WBS_ORDER || "",
                        "Sheet Material": oRow.SHEET_MATERIAL || "",
                        "Metal Grade": oRow.METAL_GRADE || "",
                        "Sheet Thickness": oRow.SHEET_THICKNESS || "",
                        "RM Length": oRow.RM_LENGTH || "",
                        "RM Width": oRow.RM_WIDTH || "",
                        "Total RM Qty": oRow.TOTAL_RM_QTY || "",
                        "A": oRow.A || "",
                        "B": oRow.B || "",
                        "C": oRow.C || "",
                        "D": oRow.D || ""
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
                { wch: 20 }, { wch: 8 }, { wch: 28 }, { wch: 18 },
                { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 12 },
                { wch: 12 }, { wch: 14 }, { wch: 6 }, { wch: 6 },
                { wch: 6 }, { wch: 6 }
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
                return oRow.MATERIAL || oRow.PLANT ||
                    oRow.STOCK_CATEGORY || oRow.WBS_ORDER ||
                    oRow.SHEET_MATERIAL || oRow.METAL_GRADE ||
                    oRow.SHEET_THICKNESS || oRow.RM_LENGTH ||
                    oRow.RM_WIDTH || oRow.TOTAL_RM_QTY;
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
        _submitBatch: function (oModel, aAllData) {
            let iSuccess = 0,
                iError = 0,
                iIndex = 0,
                iTotal = aAllData.length;

            const str = (val) => String(val ?? "").replace(/[\x00-\x1F\x7F]/g, "").trim();

            const strMax = (val, maxLen) => str(val).substring(0, maxLen);

            const dec = (val) => {
                const parsed = parseFloat(str(val));
                return isNaN(parsed) ? "0.000" : parsed.toFixed(3);
            };

            this.byId('idLineItemTable').setBusy(true);

            const fnCreateNext = () => {
                if (iIndex >= iTotal) {
                    // MessageToast.show(`${iSuccess} of ${iTotal} record(s) posted successfully!`);
                    return;
                }

                let oRow = aAllData[iIndex];

                let oPayload = {
                    "Material":       strMax(oRow.MATERIAL, 40),
                    "Plant":          strMax(oRow.PLANT, 4),
                    "StockCategory":  strMax(oRow.STOCK_CATEGORY, 20),
                    "WbsOrder":       strMax(oRow.WBS_ORDER, 24),
                    "SheetMaterial":  strMax(oRow.SHEET_MATERIAL, 4),
                    "MetalGrade":     strMax(oRow.METAL_GRADE, 20),
                    "SheetThickness": dec(oRow.SHEET_THICKNESS),
                    "RmLength":       dec(oRow.RM_LENGTH),
                    "RmWidth":        dec(oRow.RM_WIDTH),
                    "TotalRmQty":     dec(oRow.TOTAL_RM_QTY),
                    "CutA":           dec(oRow.A),
                    "CutB":           dec(oRow.B),
                    "CutC":           dec(oRow.C),
                    "CutD":           dec(oRow.D)
                };

                oModel.create("/ZCB_I_METAL", oPayload, {

                    success: () => {
                        iSuccess++;
                        iIndex++;
                        if (iIndex >= iTotal) {
                            this.byId('idLineItemTable').setBusy(false);
                            MessageToast.show(`${iSuccess} of ${iTotal} record(s) posted successfully!`);
                            this.getView().getModel("excelModel").setProperty("/results", []);
                            let oRowCount = this.byId("idRowCountText");
                            if (oRowCount) {
                                oRowCount.setText("Total Rows : 0");
                            }
                            this.byId("fileUploader").clear();

                            // this.byId("idUploadContainer").setVisible(false);
                            // this.byId("idReportContainer").setVisible(true);
                            this.byId("idReportTable").getBinding("rows").refresh();
                            return;
                        }
                        fnCreateNext();
                    },
                    error: (oError) => {
                        iError++;
                        this.byId('idLineItemTable').setBusy(false);
                        let sMsg = "";
                        try {
                            let oResp = JSON.parse(oError.responseText);
                            sMsg = oResp.error?.message?.value || oError.responseText;
                        } catch (e) {
                            sMsg = oError.responseText || oError.message || "Unknown error";
                        }

                        MessageBox.error(`Row ${iIndex + 1} failed:\n${sMsg}`);
                        iIndex++;
                        fnCreateNext();
                    }
                });
            };
            fnCreateNext();
        },
        formatDate: function (sValue) {
            if (!sValue) return "";
            let oDate = new Date(sValue);
            if (isNaN(oDate)) return sValue;
            let sDay   = String(oDate.getDate()).padStart(2, "0");
            let sMonth = String(oDate.getMonth() + 1).padStart(2, "0");
            let sYear  = oDate.getFullYear();
            return sDay + "/" + sMonth + "/" + sYear;
        },
        // formatUserName: function (sId) {
        //     if (!sId) return "";
        //     let oMap = this.getView().getModel("userModel");
        //     return (oMap && oMap.getProperty("/" + sId)) || sId;
        // },

        onReportView: function () {
            this._showContainer("idReportContainer");
            this._openMaterialFilterDialog();
        },
        onBackFromReport: function () {
            this._showContainer("idUploadContainer");
        },
        onReportRefresh: function () {
            this._openMaterialFilterDialog();
        },
        _showContainer: function (sActiveId) {
            [
                "idUploadContainer",
                "idUploadContainerPipeView",
                "idReportContainer",
                "idPipeReportContainer"
            ].forEach(id => this.byId(id).setVisible(id === sActiveId));
        },
        onPipeReport: function () {
            // this._showContainer("idUploadContainerPipeView");
            let oContainer = this.byId("idUploadContainer");
            oContainer.setBusy(true);
            setTimeout(() => {
                oContainer.setBusy(false);
                this._showContainer("idUploadContainerPipeView");
            }, 2000);
        },
        onMetalReport: function () {
            // this._showContainer("idUploadContainer");
            let oContainer = this.byId("idUploadContainerPipeView");
            oContainer.setBusy(true);
            setTimeout(() => {
                oContainer.setBusy(false);
                this._showContainer("idUploadContainer");
            }, 2000);
        },
        onReportPipeView: function () {
            this._showContainer("idPipeReportContainer");
            this._bindPipeReportTable();
        },
        onBackFromPipeReport: function () {
            this._showContainer("idUploadContainerPipeView");
        },
        _openMaterialFilterDialog: function () {
            if (!this._oMaterialDialog) {
                this._oMaterialDialog = this.loadFragment({
                    name: "mrpexceluploader.fragments.MaterialFilter"
                });
            }
            // this._oMaterialDialog.then(oDialog => oDialog.open());
            this._oMaterialDialog.then(oDialog => {
                oDialog.setModel(this.getOwnerComponent().getModel());
                oDialog.getBinding("items").filter([]);
                this.byId("idReportTable").unbindRows();
                this.getView().getModel("reportCount").setProperty("/count", 0);
                oDialog.open();
            });
        },
        onMaterialSearch: function (oEvent) {
            let sValue = oEvent.getParameter("value") || "";
            let aFilters = [];
            if (sValue) {
                aFilters.push(new sap.ui.model.Filter({
                    filters: [
                        new Filter("Product",     FilterOperator.Contains, sValue),
                        new Filter("ProductName", FilterOperator.Contains, sValue)
                    ],
                    and: false
                }));
            }
            oEvent.getSource().getBinding("items").filter(aFilters);
        },
        onMaterialConfirm: function (oEvent) {
            let oItem = oEvent.getParameter("selectedItem");
            if (oItem) {
                this._sSelectedMaterial = oItem.getTitle();
                this._bindReportTable(this._sSelectedMaterial);
            }
        },
        onMaterialCancel: function () {
            this._showContainer("idUploadContainer");
        },
        _bindReportTable: function (sMatnr) {
            let oTable = this.byId("idReportTable");
            oTable.setBusy(true);

            let sPath = "/ZCB_I_METAL_RPT(P_material='" + sMatnr + "')/Set";

            oTable.bindRows({
                path: sPath,
                parameters: {
                    operationMode: "Server"
                },
                events: {
                    dataReceived: () => {
                        oTable.setBusy(false);
                        let iCount = oTable.getBinding("rows").getLength();
                        this.getView().getModel("reportCount").setProperty("/count", iCount);
                        MessageToast.show(iCount + " record(s) founds");
                    }
                }
            });
        },
        onSendToTablePipe: function () {
            if (typeof XLSX === "undefined") {
                MessageToast.show("SheetJS (XLSX) library is not loaded.");
                return;
            }

            let oUploader = this.byId("fileUploaderPipe");
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
                    let sheet = workbook.Sheets[workbook.SheetNames[0]];
                    let excelData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

                    if (!excelData || excelData.length === 0) {
                        MessageBox.error("No data found in the Excel sheet.", { title: "Empty Sheet" });
                        return;
                    }

                    let aNonEmptyRows = excelData.filter(item =>
                        Object.values(item).some(val => val !== null && val !== undefined && String(val).trim() !== "")
                    );

                    if (aNonEmptyRows.length === 0) {
                        MessageBox.error("Excel file contains column headers only.", { title: "No Data Rows" });
                        return;
                    }

                    let aResults = aNonEmptyRows.map(item => ({
                        MATERIAL:          item["Material"]              || "",
                        PLANT:             item["Plant"]                 || "",
                        STOCK_CATEGORY:    item["Stock Category"]        || "",
                        WBS_ORDER:         item["WBS / Order"]           || "",
                        PIPE_MATERIAL:     item["PIPE material"]         || "",
                        METAL_GRADE:       item["Metal Grade"]           || "",
                        RM_SIZE:           item["RM Size"]               || "",
                        RM_UOM:            item["UOM"]                   || "",
                        CONSUME_SIZE:      item["Consume size"]          || "",
                        CONSUME_UOM:       item["UOM_1"]                 || "",
                        RMN_SIZE:          item["RMN SIZE WILL BE"]      || "",
                        RMN_UOM:           item["UOM_2"]                 || "",
                        INTERNAL_DIAMETER: item["Internal diameter"]     || "",
                        EXTERNAL_DIAMETER: item["External diameter"]     || "",
                        LENGTH:            item["Length"]                || "",
                        LENGTH_UOM:        item["Unit of measurement"]   || ""
                    }));

                    if (!this.getView().getModel("pipeExcelModel")) {
                        this.getView().setModel(new JSONModel({ results: [] }), "pipeExcelModel");
                    }
                    this.getView().getModel("pipeExcelModel").setProperty("/results", aResults);
                    MessageToast.show(aResults.length + " row(s) loaded into table.");

                } catch (err) {
                    MessageBox.error("Failed to read Excel file:\n" + err.message, { title: "Read Error" });
                }
            }.bind(this);

            oReader.onerror = () => MessageBox.error("Could not read the file.", { title: "File Error" });
            oReader.readAsBinaryString(oFile);
        },
        onDownloadTemplatePipe: function () {   
            if (typeof XLSX === "undefined") {
                MessageBox.error("SheetJS library is not loaded.", { title: "Library Missing" });
                return;
            }

            let aHeaders = [
                "Material", "Plant", "Stock Category", "WBS / Order",
                "PIPE material", "Metal Grade", "RM Size", "UOM",
                "Consume size", "UOM", "RMN SIZE WILL BE", "UOM",
                "Internal diameter", "External diameter", "Length", "Unit of measurement"
            ];

            let oEmptyRow = {};
            aHeaders.forEach(h => oEmptyRow[h] = "");

            let oWorksheet = XLSX.utils.json_to_sheet([oEmptyRow], { header: aHeaders });
            oWorksheet["!cols"] = aHeaders.map(() => ({ wch: 18 }));
            let oWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(oWorkbook, oWorksheet, "Pipe Upload Template");
            XLSX.writeFile(oWorkbook, "MRP_Pipe_Upload_Template.xlsx");
            MessageToast.show("Pipe template downloaded.");
        },
        onProcessPipe: function () {
            let aAllData = this.getView().getModel("pipeExcelModel")?.getProperty("/results");

            if (!aAllData || aAllData.length === 0) {
                MessageBox.error("No line items found.\nPlease upload an Excel file.", { title: "No Data" });
                return;
            }

            let aValidRows = aAllData.filter(oRow =>
                oRow.MATERIAL || oRow.PLANT || oRow.STOCK_CATEGORY ||
                oRow.WBS_ORDER || oRow.PIPE_MATERIAL || oRow.METAL_GRADE
            );

            if (aValidRows.length === 0) {
                MessageBox.error("No valid data rows found.", { title: "Empty Data" });
                return;
            }

            MessageBox.confirm(
                "Process all " + aValidRows.length + " pipe line item(s) to backend?",
                {
                    title: "Confirm Process",
                    actions: [MessageBox.Action.YES, MessageBox.Action.NO],
                    emphasizedAction: MessageBox.Action.YES,
                    onClose: sAction => {
                        if (sAction === MessageBox.Action.YES) {
                            this._callBackendPipe(aValidRows);
                        }
                    }
                }
            );
        },
        _callBackendPipe: function (aAllData) {
            let oModel = this.getOwnerComponent().getModel();
            if (!oModel) {
                MessageBox.error("OData model not found.", { title: "Model Error" });
                return;
            }

            jQuery.ajax({
                url: "/sap/opu/odata/sap/ZCB_SB_METAL/ZCB_I_PIPE",
                method: "GET",
                headers: { "X-CSRF-Token": "Fetch", "Accept": "application/json" },
                success: (oData, sStatus, oXHR) => {
                    let sCsrfToken = oXHR.getResponseHeader("X-CSRF-Token");
                    if (!sCsrfToken) {
                        MessageBox.error("CSRF token not returned.", { title: "Auth Error" });
                        return;
                    }
                    oModel.oHeaders["X-CSRF-Token"] = sCsrfToken;
                    this._submitBatchPipe(oModel, aAllData);
                },
                error: oXHR => MessageBox.error(
                    "CSRF token fetch failed: " + oXHR.status + " " + oXHR.statusText,
                    { title: "Auth Error" }
                )
            });
        },
        _submitBatchPipe: function (oModel, aAllData) {
            let iSuccess = 0,
                iError   = 0,
                iIndex   = 0,
                iTotal   = aAllData.length;

            const str    = val => String(val ?? "").replace(/[\x00-\x1F\x7F]/g, "").trim();
            const strMax = (val, maxLen) => str(val).substring(0, maxLen);
            const dec = (val) => {
                const parsed = parseFloat(str(val));
                return isNaN(parsed) ? "0.000" : parsed.toFixed(3);
            };

            this.byId("idLineItemTablePipe").setBusy(true);

            const fnCreateNext = () => {
                if (iIndex >= iTotal) return;

                let oRow = aAllData[iIndex];

                let oPayload = {
                    "Material":         strMax(oRow.MATERIAL,           18),  
                    "Plant":            strMax(oRow.PLANT,               4),  
                    "StockCategory":    strMax(oRow.STOCK_CATEGORY,     20),  
                    "WbsOrder":         strMax(oRow.WBS_ORDER,          24),  
                    "PipeMaterial":     strMax(oRow.PIPE_MATERIAL,      30),  
                    "MetalGrade":       strMax(oRow.METAL_GRADE,        20), 
                    "RMsize":           dec(oRow.RM_SIZE),
                    "RMUom":            strMax(oRow.RM_UOM,              3),  
                    "ConsumeSize":      dec(oRow.CONSUME_SIZE),
                    "ConsumeUom":       strMax(oRow.CONSUME_UOM,         3),  
                    "RMNsize":          dec(oRow.RMN_SIZE),
                    "RMNUom":           strMax(oRow.RMN_UOM,             3),  
                    "InternalDiameter": dec(oRow.INTERNAL_DIAMETER),
                    "ExternalDiameter": dec(oRow.EXTERNAL_DIAMETER),
                    "Length":           dec(oRow.LENGTH),
                    "LengthUom":        strMax(oRow.LENGTH_UOM,          3)   
                };

                oModel.create("/ZCB_I_PIPE", oPayload, {
                    success: () => {
                        iSuccess++;
                        iIndex++;
                        if (iIndex >= iTotal) {
                            this.byId("idLineItemTablePipe").setBusy(false);
                            MessageToast.show(`${iSuccess} of ${iTotal} pipe record(s) posted successfully!`);
                            this.getView().getModel("pipeExcelModel").setProperty("/results", []);
                            this.byId("fileUploaderPipe").clear();
                            this._showContainer("idPipeReportContainer");
                            // this._bindPipeReportTable();
                            return;
                        }
                        fnCreateNext();
                    },
                    error: oError => {
                        iError++;
                        this.byId("idLineItemTablePipe").setBusy(false);
                        let sMsg = "";
                        try {
                            let oResp = JSON.parse(oError.responseText);
                            sMsg = oResp.error?.message?.value || oError.responseText;
                        } catch (e) {
                            sMsg = oError.responseText || oError.message || "Unknown error";
                        }
                        MessageBox.error(`Row ${iIndex + 1} failed:\n${sMsg}`);
                        iIndex++;
                        fnCreateNext();
                    }
                });
            };

            
            fnCreateNext();
        },
        _bindPipeReportTable: function (sMatnr) {
            let oTable = this.byId("idPipeReportTable");
            oTable.setBusy(true);

            let sPath = "/ZCB_I_PIPE_RPT(P_material='" + sMatnr + "')/Set";

            oTable.bindRows({
                path: sPath,
                parameters: { operationMode: "Server" },
                events: {
                    dataReceived: () => {
                        oTable.setBusy(false);
                        let iCount = oTable.getBinding("rows").getLength();
                        MessageToast.show(iCount + " pipe record(s) found.");
                    }
                }
            });
        },
        onPipeReportRefresh: function () {

            // this._bindPipeReportTable();
        },
    });
});