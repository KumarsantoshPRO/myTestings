sap.ui.define(["sap/ui/core/mvc/Controller", "sap/ui/model/json/JSONModel", "sap/m/MessageBox", "sap/m/MessageToast", "sap/m/Button", "sap/m/Dialog", "sap/m/Input", "sap/m/Text", "sap/ui/core/Fragment", "sap/ui/model/Filter", "sap/ui/model/FilterOperator", "sap/ui/core/format/DateFormat", "sap/ui/core/BusyIndicator"], function (Controller, JSONModel, MessageBox, MessageToast, Button, Dialog, Input, Text, Fragment, Filter, FilterOperator, DateFormat, BusyIndicator) {
  "use strict";

  const formatINR = amount => {
    const value = typeof amount === 'string' ? parseFloat(amount) : amount;
    if (isNaN(value)) return '0.00';
    return new Intl.NumberFormat('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(value);
  };
  const View = Controller.extend("webapp.controller.View", {
    constructor: function constructor() {
      Controller.prototype.constructor.apply(this, arguments);
      this._pDialog = null;
      this.sLogoBase64 = "";
      this.sSignaturBase64 = "";
      this.sStorageKey = "my_billing_app_draft_data";
      this.sSequenceKey = "my_billing_app_invoice_seq_counter";
      // Centralized Cloud Analytics Configuration
      this.sAnalyticsBinId = "6a1015ed6877513b27b3681c";
      // Storage Keys for Custom Template Tracking
      this.sCustomNumKey = "my_billing_app_custom_numeric_seq";
      this.sCustomSuffixKey = "my_billing_app_custom_suffix_format";
      this.sCustomTriggerFlag = "my_billing_app_use_custom_start_flag";
      // CENTRALIZED CLOUD STORAGE CONFIGURATION (jsonbin.io)
      // Replace these placeholder strings with your actual keys from Step 1
      this.sBinId = "6a0ffb516610dd3ae8873764";
      this.sMasterKey = "$2a$10$QW2jbDsLe9nN3eAqSzg6v.Zz3jHv6WQfDk.HLgm4T3V0uvumxbx8i";
    },
    onInit: function _onInit() {
      let oData;
      const sSavedData = localStorage.getItem(this.sStorageKey);
      var oToday = new Date();
      var oDateFormat = DateFormat.getInstance({
        pattern: "dd-MM-yyyy"
      });
      var sTodayDate = oDateFormat.format(oToday);
      if (sSavedData) {
        try {
          oData = JSON.parse(sSavedData);
        } catch (e) {
          oData = null;
        }
      }
      if (!oData) {
        oData = {
          header: {
            To: "",
            Date: sTodayDate,
            Subject: "",
            AddtionalInfo: "",
            Notes: "",
            TermsAndConditions: "",
            BankDetails: "Payment Mode: Via Online\nBank: State Bank of India,\nBranch: Mallathahalli Branch\nName: In - Telecom Services\nC/A No: 64064045533\nIFSC Code: SBIN0040457"
          },
          products: [{
            productName: "",
            quantity: 1,
            price: "0.00",
            symbol: "",
            total: "0.00"
          }],
          taxHeader: {
            To: "",
            GSTNo: "29AGKPP7288F1Z0",
            InvoiceNo: "",
            Date: sTodayDate,
            PONo: "",
            PODate: "",
            PartyGST: "",
            BankDetails: "Payment Mode: Via Online\nBank: State Bank of India,\nBranch: Mallathahalli Branch\nName: In - Telecom Services\nC/A No: 64064045533\nIFSC Code: SBIN0040457"
          },
          taxProducts: [{
            taxpProductName: "",
            taxHSNCode: "",
            taxQuantity: 1,
            taxPrice: 0,
            taxSymbol: "",
            taxTotal: "0.00"
          }],
          cashHeader: {
            cashTo: "",
            cashDate: sTodayDate,
            cashbankDetails: "Payment Mode: Via Online\nBank: State Bank of India,\nBranch: Mallathahalli Branch\nName: In - Telecom Services\nC/A No: 64064045533\nIFSC Code: SBIN0040457"
          },
          cashProducts: [{
            cashBody: "",
            cashQuantity: "1",
            cashAmount: "0.00"
          }],
          gst: "0.00",
          totalSum: "0.00",
          taxtotal: "0.00",
          taxcgst: "0.00",
          taxsgst: "0.00",
          taxtotalSum: "0.00",
          cashTotalSum: "0.00"
        };
      }
      this.getView()?.setModel(new JSONModel(oData));
      this._loadLocalLogo("img/logo.jpg");
      this._loadSignature("img/Signature.jpg");
      this._updateSectionVisibilities("Quotation");
    },
    _getDefaultFYLabel: function _getDefaultFYLabel() {
      const oToday = new Date();
      const iCurrentMonth = oToday.getMonth();
      const iCurrentYear = oToday.getFullYear();
      let iStartFY = iCurrentMonth >= 3 ? iCurrentYear : iCurrentYear - 1;
      let iEndFY = iStartFY + 1;
      return `${iStartFY.toString().substring(2)}-${iEndFY.toString().substring(2)}`;
    },
    _getFirstLineName: function _getFirstLineName(sAddress) {
      if (!sAddress || sAddress.trim() === "") return "Document";
      return sAddress.split("\n")[0].replace(/[/\\?%*:|"<>\s]/g, "_").trim();
    },
    onGenerateNextInvoiceNumber: async function _onGenerateNextInvoiceNumber() {
      const oModel = this.getView()?.getModel();
      BusyIndicator.show(0);
      try {
        let iNextSequence;
        let sSuffixFormat;
        const sIsCustomTriggerActive = localStorage.getItem(this.sCustomTriggerFlag);
        if (sIsCustomTriggerActive === "X") {
          // A. Local dialog manual override loop
          iNextSequence = parseInt(localStorage.getItem(this.sCustomNumKey) || "1");
          sSuffixFormat = localStorage.getItem(this.sCustomSuffixKey) || `/${this._getDefaultFYLabel()}`;
          localStorage.removeItem(this.sCustomTriggerFlag);
        } else {
          // B. Fetch the global live record counter from the cloud container
          const response = await fetch(`https://api.jsonbin.io/v3/b/${this.sBinId}/latest`, {
            method: "GET",
            headers: {
              "X-Master-Key": this.sMasterKey
            }
          });
          if (!response.ok) throw new Error("Could not retrieve current sequence from cloud server.");
          const result = await response.json();
          const iCurrentSequence = parseInt(result.record.counter) || 0;
          iNextSequence = iCurrentSequence + 1;
          sSuffixFormat = localStorage.getItem(this.sCustomSuffixKey) || `/${this._getDefaultFYLabel()}`;
        }

        // C. Push the updated sequential index back up to the cloud repository synchronously
        const updateResponse = await fetch(`https://api.jsonbin.io/v3/b/${this.sBinId}`, {
          method: "PUT",
          headers: {
            "Content-Type": "application/json",
            "X-Master-Key": this.sMasterKey
          },
          body: JSON.stringify({
            counter: iNextSequence
          })
        });
        if (!updateResponse.ok) throw new Error("Failed to write updated sequence index block back to database.");

        // D. Keep local backups synchronized
        localStorage.setItem(this.sCustomNumKey, iNextSequence.toString());
        const sPaddedSequence = iNextSequence < 10 ? `0${iNextSequence}` : `${iNextSequence}`;
        const sFinalInvoiceNo = `${sPaddedSequence}${sSuffixFormat}`;
        oModel.setProperty("/taxHeader/InvoiceNo", sFinalInvoiceNo);
        MessageToast.show(`Global Invoice generated successfully: ${sFinalInvoiceNo}`);
      } catch (oError) {
        MessageBox.error(`Global Counter Error: ${oError.message || oError}`);
      } finally {
        BusyIndicator.hide();
      }
    },
    onSetGlobalInvoiceSequence: function _onSetGlobalInvoiceSequence() {
      const oModel = this.getView()?.getModel();
      const oInput = new Input({
        placeholder: "e.g., 99/26-27",
        width: "100%"
      });
      const oDialog = new Dialog({
        title: "Set Global Invoice Sequence Start Template",
        type: "Message",
        content: [new Text({
          text: "Enter the custom starting sequence template matching your format (e.g., 99/26-27):"
        }), oInput],
        beginButton: new Button({
          text: "OK",
          press: () => {
            const sRawInput = oInput.getValue() ? oInput.getValue().trim() : "";
            if (!sRawInput) {
              MessageBox.error("Invalid entry. Input field cannot be empty.");
              return;
            }
            const aParts = sRawInput.split("/");
            const sNumericPart = aParts[0].trim();
            const iNewSequenceValue = parseInt(sNumericPart);
            if (isNaN(iNewSequenceValue) || iNewSequenceValue < 0) {
              MessageBox.error("Invalid template number sequence. Ensure it begins with a valid whole number.");
              return;
            }
            const sExtractedSuffix = aParts.length > 1 ? `/${aParts.slice(1).join("/")}` : `/${this._getDefaultFYLabel()}`;

            // Synchronously reset the centralized cloud database tracker value
            BusyIndicator.show(0);
            fetch(`https://api.jsonbin.io/v3/b/${this.sBinId}`, {
              method: "PUT",
              headers: {
                "Content-Type": "application/json",
                "X-Master-Key": this.sMasterKey
              },
              body: JSON.stringify({
                counter: iNewSequenceValue
              })
            }).then(res => {
              if (!res.ok) throw new Error("Database update rejected.");
              localStorage.setItem(this.sCustomNumKey, iNewSequenceValue.toString());
              localStorage.setItem(this.sCustomSuffixKey, sExtractedSuffix);
              localStorage.setItem(this.sCustomTriggerFlag, "X");
              oModel.setProperty("/taxHeader/InvoiceNo", "");
              MessageToast.show(`Global sequence reset successfully to: ${sRawInput}`);
              oDialog.close();
            }).catch(err => {
              MessageBox.error(`Cloud Sync Failed: ${err.message}`);
            }).finally(() => {
              BusyIndicator.hide();
            });
          }
        }),
        endButton: new Button({
          text: "Cancel",
          press: () => {
            oDialog.close();
          }
        }),
        afterClose: () => {
          oDialog.destroy();
        }
      });
      oDialog.open();
    },
    onSaveLocalStorage: function _onSaveLocalStorage() {
      const oModel = this.getView()?.getModel();
      if (oModel) {
        localStorage.setItem(this.sStorageKey, JSON.stringify(oModel.getData()));
        MessageToast.show("Form data successfully saved locally!");
      } else {
        MessageBox.error("Error updating model context details.");
      }
    },
    onClearLocalStorage: function _onClearLocalStorage() {
      MessageBox.confirm("Are you sure you want to clear your saved draft?", {
        actions: [MessageBox.Action.YES, MessageBox.Action.NO],
        onClose: sAction => {
          if (sAction === MessageBox.Action.YES) {
            localStorage.removeItem(this.sStorageKey);
            MessageToast.show("Saved data deleted. Refreshing page layout.");
            this.onInit();
          }
        }
      });
    },
    _loadLocalLogo: function _loadLocalLogo(sRelativePath) {
      const sFullUrl = sap.ui.require.toUrl("my/app/generatebill/" + sRelativePath);
      const xhr = new XMLHttpRequest();
      xhr.onload = () => {
        const reader = new FileReader();
        reader.onloadend = () => {
          this.sLogoBase64 = reader.result;
        };
        reader.readAsDataURL(xhr.response);
      };
      xhr.open("GET", sFullUrl);
      xhr.responseType = "blob";
      xhr.send();
    },
    _loadSignature: function _loadSignature(sSigRelativePath) {
      const sFullUrl = sap.ui.require.toUrl("my/app/generatebill/" + sSigRelativePath);
      const xhr = new XMLHttpRequest();
      xhr.onload = () => {
        const reader = new FileReader();
        reader.onloadend = () => {
          this.sSignaturBase64 = reader.result;
        };
        reader.readAsDataURL(xhr.response);
      };
      xhr.open("GET", sFullUrl);
      xhr.responseType = "blob";
      xhr.send();
    },
    onSelectChange: function _onSelectChange(oEvent) {
      const iSelectedText = oEvent.getParameter("selectedItem").getText();
      this._updateSectionVisibilities(iSelectedText);
    },
    _updateSectionVisibilities: function _updateSectionVisibilities(sMode) {
      const view = this.getView();
      if (!view) return;
      view.byId("idOPSQuote1").setVisible(sMode === "Quotation");
      view.byId("idOPSQuote2").setVisible(sMode === "Quotation");
      view.byId("idQuotationSec").setVisible(sMode === "Quotation");
      view.byId("idBtnQuotation").setVisible(sMode === "Quotation");
      view.byId("chkBankDetail").setVisible(sMode === "Quotation");
      view.byId("chkGST").setVisible(sMode === "Quotation");
      view.byId("chkTotal").setVisible(sMode === "Quotation");
      view.byId("idOPSQuoteTaxInc").setVisible(sMode === "TAX-INVOICE");
      view.byId("idTaxSec").setVisible(sMode === "TAX-INVOICE");
      view.byId("idBtnInvoice").setVisible(sMode === "TAX-INVOICE");
      view.byId("idOPSCash").setVisible(sMode === "Cash Bill");
      view.byId("idCashSecTab").setVisible(sMode === "Cash Bill");
      view.byId("idBtnCash").setVisible(sMode === "Cash Bill");
    },
    onAddRow: function _onAddRow() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/products");
      aProducts.push({
        productName: "",
        price: "0.00",
        symbol: "",
        quantity: 1,
        total: "0.00"
      });
      oModel.setProperty("/products", aProducts);
    },
    onDelete: function _onDelete(oEvent) {
      const oItemToDelete = oEvent.getParameter("listItem");
      const sPath = oItemToDelete.getBindingContext().getPath();
      const iIndex = parseInt(sPath.split("/").pop());
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/products");
      aProducts.splice(iIndex, 1);
      oModel.setProperty("/products", aProducts);
      this.onCalc();
    },
    onCalc: function _onCalc() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/products");
      let fGrandTotal = 0;
      let gstAmount = 0;
      aProducts.forEach(oProduct => {
        const fQty = parseFloat(oProduct.quantity) || 0;
        const fPrice = parseFloat(oProduct.price) || 0;
        let fRowTotal = fQty * fPrice;
        oProduct.total = fRowTotal.toFixed(2);
        gstAmount += fRowTotal * 0.18;
        fGrandTotal += fRowTotal;
      });
      oModel.setProperty("/gst", gstAmount.toFixed(2));
      oModel.setProperty("/totalSum", (fGrandTotal + gstAmount).toFixed(2));
      oModel.refresh();
    },
    onGeneratePDF: function _onGeneratePDF() {
      this.onSaveLocalStorage();
      const jspdfLib = window.jspdf;
      if (!jspdfLib) return;
      const oModel = this.getView()?.getModel();
      const oHeader = oModel.getProperty("/header");
      const aItems = oModel.getProperty("/products");
      const doc = new jspdfLib.jsPDF();
      const pageWidth = doc.internal.pageSize.width;
      const pageHeight = doc.internal.pageSize.height;
      const margin = 5;
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.3);
      doc.rect(margin, margin, pageWidth - margin * 2, pageHeight - margin * 2);
      if (this.sLogoBase64) {
        doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25);
      }
      doc.setFontSize(10);
      doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, {
        align: 'right'
      });
      doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, {
        align: 'right'
      });
      doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, {
        align: 'right'
      });
      doc.setFontSize(9);
      doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, {
        align: 'right'
      });
      doc.text("Nagarabhavi, Bangalore-560072", pageWidth - 14, 35, {
        align: 'right'
      });
      doc.line(5, 40, pageWidth - 5, 40);
      doc.setFont("helvetica", "bold");
      doc.text(`Date: ${oHeader.Date}`, pageWidth - 14, 45, {
        align: 'right'
      });
      doc.text("To,", 14, 45);
      doc.setFont("helvetica", "normal");
      doc.text(doc.splitTextToSize(oHeader.To, 80), 14, 50);
      doc.setFont("helvetica", "bold");
      let finalHeaderY = 70;
      doc.text("Sub: " + oHeader.Subject, 14, 70);
      doc.setFont("helvetica", "normal");
      if (oHeader.AddtionalInfo !== "") {
        doc.text(oHeader.AddtionalInfo, 14, finalHeaderY + 4);
      }
      const bShowGST = (this.getView()?.byId("chkGST")).getSelected();
      const bShowTotal = (this.getView()?.byId("chkTotal")).getSelected();
      const tableBody = aItems.map((item, index) => [index + 1, item.productName, item.quantity, item.price + item.symbol, formatINR(item.total)]);
      const subtotal = aItems.reduce((acc, cur) => acc + parseFloat(cur.total || 0), 0);
      const gstAmount = subtotal * 0.18;
      const grandTotal = subtotal + gstAmount;
      if (bShowGST) {
        tableBody.push([{
          content: '18% GST Amount',
          colSpan: 4,
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }, {
          content: formatINR(gstAmount),
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }]);
      }
      if (bShowTotal) {
        tableBody.push([{
          content: 'Total',
          colSpan: 4,
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }, {
          content: formatINR(grandTotal),
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }]);
      }
      doc.autoTable({
        startY: finalHeaderY + 8,
        head: [['Sl.No.', 'Particulars', 'Quantity', 'Rate', 'Total (Rs.)']],
        body: tableBody,
        theme: 'grid',
        styles: {
          fontSize: 9,
          lineColor: [0, 0, 0],
          lineWidth: 0.1
        },
        headStyles: {
          fillColor: [255, 255, 255],
          textColor: [0, 0, 0],
          fontStyle: 'bold'
        },
        columnStyles: {
          0: {
            cellWidth: 15
          },
          1: {
            cellWidth: 'auto'
          },
          2: {
            cellWidth: 20,
            halign: 'center'
          },
          3: {
            cellWidth: 25,
            halign: 'right'
          },
          4: {
            cellWidth: 30,
            halign: 'right'
          }
        }
      });
      let finalY = doc.lastAutoTable.finalY;
      if (oHeader.TermsAndConditions !== "") {
        finalY += 10;
        doc.setFont("helvetica", "bold");
        doc.text("Terms & Conditions:", 14, finalY);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.TermsAndConditions, pageWidth - 28), 14, finalY + 4);
      }
      if (oHeader.Notes !== "") {
        finalY += 20;
        doc.setFont("helvetica", "bold");
        doc.text("Notes:", 14, finalY);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.Notes, pageWidth - 28), 14, finalY + 4);
      }
      if ((this.getView()?.byId("chkBankDetail")).getSelected()) {
        finalY += 20;
        doc.setFont("helvetica", "bold");
        doc.text("Bank Details:", 14, finalY);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.BankDetails, pageWidth - 28), 14, finalY + 4);
      }
      if (this.sSignaturBase64) {
        doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, pageHeight - 40, 70, 25);
      }
      const sTargetFilename = this._getFirstLineName(oHeader.To) + "_Quotation.pdf";
      doc.save(sTargetFilename);
      this.onExcelDownload(null, "Quotation");
      this.onGenerateWord("Quotation");
      const sClient = this._getFirstLineName(oHeader.To);
      this._logDocumentAnalytics("Quotation", sTargetFilename, sClient, grandTotal);
    },
    onTaxAddRow: function _onTaxAddRow() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/taxProducts");
      aProducts.push({
        taxpProductName: "",
        taxHSNCode: "",
        taxQuantity: 1,
        taxPrice: 0,
        taxSymbol: "",
        taxTotal: "0.00"
      });
      oModel.setProperty("/taxProducts", aProducts);
    },
    onTaxDelete: function _onTaxDelete(oEvent) {
      const oItemToDelete = oEvent.getParameter("listItem");
      const sPath = oItemToDelete.getBindingContext().getPath();
      const iIndex = parseInt(sPath.split("/").pop());
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/taxProducts");
      aProducts.splice(iIndex, 1);
      oModel.setProperty("/taxProducts", aProducts);
      this.onTaxCalc();
    },
    onTaxCalc: function _onTaxCalc() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/taxProducts");
      let fGrandTotal = 0;
      let taxcgst = 0;
      let taxsgst = 0;
      aProducts.forEach(oProduct => {
        const fQty = parseFloat(oProduct.taxQuantity) || 0;
        const fPrice = parseFloat(oProduct.taxPrice) || 0;
        let fRowTotal = fQty * fPrice;
        oProduct.taxTotal = fRowTotal.toFixed(2);
        taxcgst += fRowTotal * 0.09;
        taxsgst += fRowTotal * 0.09;
        fGrandTotal += fRowTotal;
      });
      oModel.setProperty("/taxtotal", fGrandTotal.toFixed(2));
      oModel.setProperty("/taxcgst", taxcgst.toFixed(2));
      oModel.setProperty("/taxsgst", taxsgst.toFixed(2));
      oModel.setProperty("/taxtotalSum", (fGrandTotal + taxcgst + taxsgst).toFixed(2));
      oModel.refresh();
    },
    onTaxInvoicePDF: function _onTaxInvoicePDF() {
      this.onSaveLocalStorage();
      const jspdfLib = window.jspdf;
      if (!jspdfLib) return;
      const oModel = this.getView()?.getModel();
      const oData = oModel.getData();
      const doc = new jspdfLib.jsPDF();
      const startX = 5;
      const endX = 205;
      const pageWidth = doc.internal.pageSize.width;
      const pageHeight = doc.internal.pageSize.height;
      const margin = 5;
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.3);
      doc.rect(margin, margin, pageWidth - margin * 2, pageHeight - margin * 2);
      if (this.sLogoBase64) {
        doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25);
      }
      doc.setFontSize(10);
      doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, {
        align: 'right'
      });
      doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, {
        align: 'right'
      });
      doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, {
        align: 'right'
      });
      doc.setFontSize(9);
      doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, {
        align: 'right'
      });
      doc.text("Nagarabhavi, Bangalore-560072", pageWidth - 14, 35, {
        align: 'right'
      });
      doc.line(5, 40, pageWidth - 5, 40);
      doc.setFont("helvetica", "bold");
      doc.text("To,", 15, 44);
      doc.setFont("helvetica", "normal");
      doc.text(doc.splitTextToSize(oData.taxHeader.To, 90), 15, 49);
      const metaX = 130;
      doc.text(`GST No: ${oData.taxHeader.GSTNo}`, metaX, 44);
      doc.text(`Invoice No: ${oData.taxHeader.InvoiceNo}`, metaX, 50);
      doc.text(`Date: ${oData.taxHeader.Date}`, metaX, 56);
      doc.text(`PO No: ${oData.taxHeader.PONo}`, metaX, 62);
      doc.text(`PO Date: ${oData.taxHeader.PODate}`, metaX, 68);
      const tableRows = oData.taxProducts.map((item, index) => [index + 1, item.taxpProductName, item.taxHSNCode, formatINR(item.taxPrice) + item.taxSymbol, item.taxQuantity, formatINR(item.taxTotal)]);
      const totalAmount = oData.taxProducts.reduce((sum, item) => sum + parseFloat(item.taxTotal), 0);
      const taxVal = totalAmount * 0.09;
      const grandTotal = totalAmount + taxVal * 2;
      tableRows.push([{
        content: `TOTAL:`,
        colSpan: 5,
        styles: {
          halign: 'right',
          fontStyle: 'bold'
        }
      }, {
        content: formatINR(totalAmount),
        styles: {
          halign: 'right',
          fontStyle: 'bold'
        }
      }], [{
        content: `CGST @ 9%:`,
        colSpan: 5,
        styles: {
          halign: 'right'
        }
      }, {
        content: formatINR(taxVal),
        styles: {
          halign: 'right'
        }
      }], [{
        content: `SGST @ 9%:`,
        colSpan: 5,
        styles: {
          halign: 'right'
        }
      }, {
        content: formatINR(taxVal),
        styles: {
          halign: 'right'
        }
      }], [{
        content: `GRAND TOTAL:`,
        colSpan: 5,
        styles: {
          halign: 'right',
          fontStyle: 'bold',
          fontSize: 10
        }
      }, {
        content: formatINR(grandTotal),
        styles: {
          halign: 'right',
          fontStyle: 'bold',
          fontSize: 10
        }
      }]);
      doc.autoTable({
        startY: 75,
        head: [["SI No.", "Particulars", "HSN Code", "Rate", "No.of Units", "Amount"]],
        body: tableRows,
        theme: 'grid',
        styles: {
          fontSize: 8,
          lineColor: [0, 0, 0],
          lineWidth: 0.1,
          textColor: [0, 0, 0]
        },
        headStyles: {
          fillColor: [255, 255, 255],
          textColor: [0, 0, 0],
          fontStyle: 'bold',
          halign: 'center'
        },
        columnStyles: {
          0: {
            halign: 'center'
          },
          3: {
            halign: 'right'
          },
          4: {
            halign: 'center'
          },
          5: {
            halign: 'right'
          }
        },
        didParseCell: function (data) {
          const totalRowsCount = 4;
          const isTotalRow = data.row.index >= tableRows.length - totalRowsCount;
          if (isTotalRow) {
            data.cell.styles.lineWidth = {
              top: 0.1,
              right: 0,
              bottom: 0.1,
              left: 0
            };
            if (data.column.index === 5) {
              data.cell.styles.lineWidth = {
                top: 0.1,
                right: 0.1,
                bottom: 0.1,
                left: 0
              };
            }
            if (data.column.index === 0) {
              data.cell.styles.lineWidth = {
                top: 0.1,
                right: 0,
                bottom: 0.1,
                left: 0.1
              };
            }
          }
        }
      });
      const finalY = doc.lastAutoTable.finalY + 5;
      doc.line(startX, finalY + 5, endX, finalY + 5);
      doc.setFont("helvetica", "normal");
      doc.text(`Rupees in words: ${this.numberToWords(grandTotal)} Only`, 10, finalY + 10);
      doc.text(`Party GST No: ${oData.taxHeader.PartyGST || ""}`, 10, finalY + 18);
      doc.setFont("helvetica", "bold");
      doc.text("Bank Details", 10, finalY + 28);
      doc.setFont("helvetica", "normal");
      doc.text(doc.splitTextToSize(oData.taxHeader.BankDetails, 100), 10, finalY + 33);
      if (this.sSignaturBase64) {
        doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, pageHeight - 40, 70, 25);
      }
      const sTargetFilename = this._getFirstLineName(oData.taxHeader.To) + "_TaxInvoice.pdf";
      doc.save(sTargetFilename);
      this.onExcelDownload(null, "TAX-INVOICE");
      this.onGenerateWord("TAX-INVOICE");
      const sClient = this._getFirstLineName(oData.taxHeader.To);
      this._logDocumentAnalytics("TAX-INVOICE", oData.taxHeader.InvoiceNo, sClient, grandTotal);
    },
    onCashAddRow: function _onCashAddRow() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/cashProducts");
      aProducts.push({
        cashBody: "",
        cashQuantity: "1",
        cashAmount: "0.00"
      });
      oModel.setProperty("/cashProducts", aProducts);
    },
    onCashDelete: function _onCashDelete(oEvent) {
      const oItemToDelete = oEvent.getParameter("listItem");
      const sPath = oItemToDelete.getBindingContext().getPath();
      const iIndex = parseInt(sPath.split("/").pop());
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/cashProducts");
      aProducts.splice(iIndex, 1);
      oModel.setProperty("/cashProducts", aProducts);
      this.onCashCalc();
    },
    onCashCalc: function _onCashCalc() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/cashProducts") || [];
      let totalAmount = 0;
      aProducts.forEach(item => {
        totalAmount += parseFloat(item.cashAmount || 0);
      });
      oModel.setProperty("/cashTotalSum", totalAmount.toFixed(2));
    },
    onCashBillPDF: function _onCashBillPDF() {
      this.onSaveLocalStorage();
      const jspdfLib = window.jspdf;
      if (!jspdfLib) return;
      const oModel = this.getView()?.getModel();
      const oData = oModel.getData();
      const doc = new jspdfLib.jsPDF();
      const startX = 5;
      const endX = 205;
      const pageWidth = doc.internal.pageSize.width;
      const pageHeight = doc.internal.pageSize.height;
      const margin = 5;
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.3);
      doc.rect(margin, margin, pageWidth - margin * 2, pageHeight - margin * 2);
      if (this.sLogoBase64) {
        doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25);
      }
      doc.setFontSize(10);
      doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, {
        align: 'right'
      });
      doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, {
        align: 'right'
      });
      doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, {
        align: 'right'
      });
      doc.setFontSize(9);
      doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, {
        align: 'right'
      });
      doc.text("Nagarabhavi, Bangalore-560072", pageWidth - 14, 35, {
        align: 'right'
      });
      doc.line(5, 40, pageWidth - 5, 40);
      doc.setFont("helvetica", "bold");
      doc.text("To,", 15, 46);
      doc.setFont("helvetica", "normal");
      doc.text(doc.splitTextToSize(oData.cashHeader.cashTo, 90), 15, 51);
      doc.text(`Date: ${oData.cashHeader.cashDate}`, pageWidth - 14, 46, {
        align: 'right'
      });
      doc.setFont("helvetica", "bold");
      doc.setFontSize(14);
      doc.text(`Cash Bill`, pageWidth / 2, 72, {
        align: 'center'
      });
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10);
      const tableRows = oData.cashProducts.map((item, index) => [index + 1, item.cashBody, item.cashQuantity, formatINR(item.cashAmount)]);
      const totalAmount = oData.cashProducts.reduce((sum, item) => sum + parseFloat(item.cashAmount || 0), 0);
      tableRows.push([{
        content: `GRAND TOTAL:`,
        colSpan: 3,
        styles: {
          halign: 'right',
          fontStyle: 'bold',
          fontSize: 10
        }
      }, {
        content: formatINR(totalAmount),
        styles: {
          halign: 'right',
          fontStyle: 'bold',
          fontSize: 10
        }
      }]);
      doc.autoTable({
        startY: 75,
        head: [["SI No.", "Particulars", "Quantity", "Amount"]],
        body: tableRows,
        theme: 'grid',
        columnStyles: {
          0: {
            cellWidth: 15
          },
          1: {
            cellWidth: 'auto'
          },
          2: {
            cellWidth: 25,
            halign: 'center'
          },
          3: {
            cellWidth: 35,
            halign: 'right'
          }
        }
      });
      const finalY = doc.lastAutoTable.finalY + 5;
      doc.line(startX, finalY + 5, endX, finalY + 5);
      doc.setFont("helvetica", "normal");
      doc.text(`Rupees in words: ${this.numberToWords(totalAmount)} Only`, 10, finalY + 10);
      doc.setFont("helvetica", "bold");
      doc.text("Bank Details", 10, finalY + 28);
      doc.setFont("helvetica", "normal");
      doc.text(doc.splitTextToSize(oData.cashHeader.cashbankDetails, 100), 10, finalY + 33);
      if (this.sSignaturBase64) {
        doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, pageHeight - 40, 70, 25);
      }
      const sTargetFilename = this._getFirstLineName(oData.cashHeader.cashTo) + "_CashBill.pdf";
      doc.save(sTargetFilename);
      this.onExcelDownload(null, "Cash Bill");
      this.onGenerateWord("Cash Bill");
      const sClient = this._getFirstLineName(oData.cashHeader.cashTo);
      this._logDocumentAnalytics("Cash Bill", sTargetFilename, sClient, totalAmount);
    },
    onExcelDownload: function _onExcelDownload(oEvent, sOverrideMode) {
      const sSelectMode = sOverrideMode || (this.getView()?.byId("mySelect")).getSelectedItem()?.getText();
      const oModel = this.getView()?.getModel();
      let sFileName = "Export.xlsx";
      let headerData = [];
      let itemsData = [];
      headerData.push({
        "FIELD": "MODE_METADATA",
        "VALUE": sSelectMode
      });
      if (sSelectMode === "Quotation") {
        sFileName = this._getFirstLineName(oModel.getProperty("/header/To")) + "_QuotationItems.xlsx";
        headerData.push({
          "FIELD": "Header_To",
          "VALUE": oModel.getProperty("/header/To")
        });
        headerData.push({
          "FIELD": "Header_Date",
          "VALUE": oModel.getProperty("/header/Date")
        });
        headerData.push({
          "FIELD": "Header_Subject",
          "VALUE": oModel.getProperty("/header/Subject")
        });
        headerData.push({
          "FIELD": "Header_AddtionalInfo",
          "VALUE": oModel.getProperty("/header/AddtionalInfo")
        });
        headerData.push({
          "FIELD": "Header_Notes",
          "VALUE": oModel.getProperty("/header/Notes")
        });
        headerData.push({
          "FIELD": "Header_TermsAndConditions",
          "VALUE": oModel.getProperty("/header/TermsAndConditions")
        });
        headerData.push({
          "FIELD": "Header_BankDetails",
          "VALUE": oModel.getProperty("/header/BankDetails")
        });
        itemsData = oModel.getProperty("/products").map(item => ({
          "Particulars": item.productName,
          "Rate": item.price,
          "Quantity": item.quantity,
          "Symbol": item.symbol,
          "Total": item.total
        }));
      } else if (sSelectMode === "TAX-INVOICE") {
        sFileName = this._getFirstLineName(oModel.getProperty("/taxHeader/To")) + "_TaxInvoiceItems.xlsx";
        headerData.push({
          "FIELD": "TaxHeader_To",
          "VALUE": oModel.getProperty("/taxHeader/To")
        });
        headerData.push({
          "FIELD": "TaxHeader_GSTNo",
          "VALUE": oModel.getProperty("/taxHeader/GSTNo")
        });
        headerData.push({
          "FIELD": "TaxHeader_InvoiceNo",
          "VALUE": oModel.getProperty("/taxHeader/InvoiceNo")
        });
        headerData.push({
          "FIELD": "TaxHeader_Date",
          "VALUE": oModel.getProperty("/taxHeader/Date")
        });
        headerData.push({
          "FIELD": "TaxHeader_PONo",
          "VALUE": oModel.getProperty("/taxHeader/PONo")
        });
        headerData.push({
          "FIELD": "TaxHeader_PODate",
          "VALUE": oModel.getProperty("/taxHeader/PODate")
        });
        headerData.push({
          "FIELD": "TaxHeader_PartyGST",
          "VALUE": oModel.getProperty("/taxHeader/PartyGST")
        });
        itemsData = oModel.getProperty("/taxProducts").map(item => ({
          "Particulars": item.taxpProductName,
          "HSN Code": item.taxHSNCode,
          "Rate": item.taxPrice,
          "Quantity": item.taxQuantity,
          "Symbol": item.taxSymbol,
          "Total": item.taxTotal
        }));
      } else {
        sFileName = this._getFirstLineName(oModel.getProperty("/cashHeader/cashTo")) + "_CashBillItems.xlsx";
        headerData.push({
          "FIELD": "CashHeader_cashTo",
          "VALUE": oModel.getProperty("/cashHeader/cashTo")
        });
        headerData.push({
          "FIELD": "CashHeader_cashDate",
          "VALUE": oModel.getProperty("/cashHeader/cashDate")
        });
        headerData.push({
          "FIELD": "CashHeader_cashbankDetails",
          "VALUE": oModel.getProperty("/cashHeader/cashbankDetails")
        });
        itemsData = oModel.getProperty("/cashProducts").map(item => ({
          "Particulars": item.cashBody,
          "Quantity": item.cashQuantity,
          "Amount": item.cashAmount
        }));
      }
      const ws = XLSX.utils.json_to_sheet(headerData);
      XLSX.utils.sheet_add_aoa(ws, [[]], {
        origin: -1
      });
      XLSX.utils.sheet_add_json(ws, itemsData, {
        origin: -1
      });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Items Data");
      XLSX.writeFile(wb, sFileName);
    },
    onExcelUpload: function _onExcelUpload(oEvent) {
      const oFileUploader = oEvent.getSource();
      const oFile = oEvent.getParameter("files")[0];
      if (!oFile) return;
      const reader = new FileReader();
      reader.onload = e => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {
          type: 'array'
        });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rawJsonRows = XLSX.utils.sheet_to_json(worksheet, {
          header: 1
        });
        const sSelectMode = (this.getView()?.byId("mySelect")).getSelectedItem()?.getText();
        const oModel = this.getView()?.getModel();
        const headerMap = {};
        let lineItemsStartIndex = 0;
        for (let i = 0; i < rawJsonRows.length; i++) {
          const row = rawJsonRows[i];
          if (row && row[0] && typeof row[0] === 'string' && (row[0].includes("Header") || row[0] === "FIELD" || row[0] === "MODE_METADATA")) {
            if (row[0] !== "FIELD") headerMap[row[0]] = row[1] || "";
          } else {
            lineItemsStartIndex = i;
            break;
          }
        }
        const lineItemsRows = rawJsonRows.slice(lineItemsStartIndex).filter(row => row && row.length > 0);
        if (lineItemsRows.length === 0) return;
        const tableHeaders = lineItemsRows[0];
        const jsonDataItems = lineItemsRows.slice(1).map(row => {
          const obj = {};
          tableHeaders.forEach((headerName, index) => {
            obj[headerName] = row[index] !== undefined ? row[index] : "";
          });
          return obj;
        });
        if (sSelectMode === "Quotation") {
          if (headerMap["Header_To"]) oModel.setProperty("/header/To", headerMap["Header_To"]);
          if (headerMap["Header_Date"]) oModel.setProperty("/header/Date", headerMap["Header_Date"]);
          if (headerMap["Header_Subject"]) oModel.setProperty("/header/Subject", headerMap["Header_Subject"]);
          if (headerMap["Header_AddtionalInfo"]) oModel.setProperty("/header/AddtionalInfo", headerMap["Header_AddtionalInfo"]);
          if (headerMap["Header_Notes"]) oModel.setProperty("/header/Notes", headerMap["Header_Notes"]);
          if (headerMap["Header_TermsAndConditions"]) oModel.setProperty("/header/TermsAndConditions", headerMap["Header_TermsAndConditions"]);
          if (headerMap["Header_BankDetails"]) oModel.setProperty("/header/BankDetails", headerMap["Header_BankDetails"]);
          const parsedProducts = jsonDataItems.map(row => ({
            productName: row["Particulars"] || "",
            price: row["Rate"]?.toString() || "0.00",
            quantity: parseInt(row["Quantity"]) || 1,
            symbol: row["Symbol"] || "",
            total: row["Total"]?.toString() || "0.00"
          }));
          oModel.setProperty("/products", parsedProducts);
          this.onCalc();
        } else if (sSelectMode === "TAX-INVOICE") {
          if (headerMap["TaxHeader_To"]) oModel.setProperty("/taxHeader/To", headerMap["TaxHeader_To"]);
          if (headerMap["TaxHeader_GSTNo"]) oModel.setProperty("/taxHeader/GSTNo", headerMap["TaxHeader_GSTNo"]);
          if (headerMap["TaxHeader_InvoiceNo"]) oModel.setProperty("/taxHeader/InvoiceNo", headerMap["TaxHeader_InvoiceNo"]);
          if (headerMap["TaxHeader_Date"]) oModel.setProperty("/taxHeader/Date", headerMap["TaxHeader_Date"]);
          if (headerMap["TaxHeader_PONo"]) oModel.setProperty("/taxHeader/PONo", headerMap["TaxHeader_PONo"]);
          if (headerMap["TaxHeader_PODate"]) oModel.setProperty("/taxHeader/PODate", headerMap["TaxHeader_PODate"]);
          if (headerMap["TaxHeader_PartyGST"]) oModel.setProperty("/taxHeader/PartyGST", headerMap["TaxHeader_PartyGST"]);
          const parsedTax = jsonDataItems.map(row => ({
            taxpProductName: row["Particulars"] || "",
            taxHSNCode: row["HSN Code"]?.toString() || "",
            taxQuantity: parseInt(row["Quantity"]) || 1,
            taxPrice: parseFloat(row["Rate"]) || 0,
            taxSymbol: row["Symbol"] || "",
            taxTotal: row["Total"]?.toString() || "0.00"
          }));
          oModel.setProperty("/taxProducts", parsedTax);
          this.onTaxCalc();
        } else {
          if (headerMap["CashHeader_cashTo"]) oModel.setProperty("/cashHeader/cashTo", headerMap["CashHeader_cashTo"]);
          if (headerMap["CashHeader_cashDate"]) oModel.setProperty("/cashHeader/cashDate", headerMap["CashHeader_cashDate"]);
          if (headerMap["CashHeader_cashbankDetails"]) oModel.setProperty("/cashHeader/cashbankDetails", headerMap["CashHeader_cashbankDetails"]);
          const parsedCash = jsonDataItems.map(row => ({
            cashBody: row["Particulars"] || "",
            cashQuantity: row["Quantity"]?.toString() || "1",
            cashAmount: row["Amount"]?.toString() || "0.00"
          }));
          oModel.setProperty("/cashProducts", parsedCash);
          this.onCashCalc();
        }
        MessageToast.show("Excel headers and rows processed successfully!");
        oFileUploader.clear();
      };
      reader.readAsArrayBuffer(oFile);
    },
    numberToWords: function _numberToWords(num) {
      const a = ['', 'One ', 'Two ', 'Three ', 'Four ', 'Five ', 'Six ', 'Seven ', 'Eight ', 'Nine ', 'Ten ', 'Eleven ', 'Twelve ', 'Thirteen ', 'Fourteen ', 'Fifteen ', 'Sixteen ', 'Seventeen ', 'Eighteen ', 'Nineteen '];
      const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
      if (num === 0) return 'Zero';
      const convertWholeNumber = amountStr => {
        const n = ('000000000' + amountStr).substring(amountStr.length + 9 - 9).match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
        if (!n) return '';
        let str = '';
        str += Number(n[1]) !== 0 ? (a[Number(n[1])] || b[Number(n[1][0])] + ' ' + a[Number(n[1][1])]) + 'Crore ' : '';
        str += Number(n[2]) !== 0 ? (a[Number(n[2])] || b[Number(n[2][0])] + ' ' + a[Number(n[2][1])]) + 'Lakh ' : '';
        str += Number(n[3]) !== 0 ? (a[Number(n[3])] || b[Number(n[3][0])] + ' ' + a[Number(n[3][1])]) + 'Thousand ' : '';
        str += Number(n[4]) !== 0 ? (a[Number(n[4])] || b[Number(n[4][0])] + ' ' + a[Number(n[4][1])]) + 'Hundred ' : '';
        str += Number(n[5]) !== 0 ? (str !== '' ? 'and ' : '') + (a[Number(n[5])] || b[Number(n[5][0])] + ' ' + a[Number(n[5][1])]) : '';
        return str.trim();
      };
      const fixedNumStr = num.toFixed(2);
      const parts = fixedNumStr.split('.');
      const rupeePart = parts[0];
      const paisaPart = parts[1];
      let result = '';
      if (Number(rupeePart) > 0) {
        result += convertWholeNumber(rupeePart);
      } else if (Number(paisaPart) > 0) {
        result += 'Zero';
      }
      if (paisaPart && Number(paisaPart) > 0) {
        const p = paisaPart.match(/^(\d{2})$/);
        if (p) {
          const paisaWords = a[Number(p[1])] || b[Number(p[1][0])] + ' ' + a[Number(p[1][1])];
          result += (result !== '' ? ' and ' : '') + paisaWords.trim() + ' Paise';
        }
      }
      return result.trim();
    },
    onHSNValueHelpRequest: async function _onHSNValueHelpRequest(oEvent) {
      const oInput = oEvent.getSource();
      this.getView()?.data("valueHelpSourceInput", oInput);
      if (!this._pDialog) {
        this._pDialog = Fragment.load({
          id: this.getView()?.getId(),
          name: "my.app.generatebill.view.HSNValueHelpDialog",
          controller: this
        });
        const oDialog = await this._pDialog;
        this.getView()?.addDependent(oDialog);
      }
      const oDialog = await this._pDialog;
      const oBinding = oDialog.getBinding("items");
      if (oBinding) oBinding.filter([]);
      oDialog.open("");
    },
    onHSNValueHelpSearch: function _onHSNValueHelpSearch(oEvent) {
      const sValue = oEvent.getParameter("value");
      const oDialog = oEvent.getSource();
      const oBinding = oDialog.getBinding("items");
      if (!sValue) {
        oBinding.filter([]);
        return;
      }
      const oFilterCode = new Filter("HSN_CD", FilterOperator.Contains, sValue);
      const oFilterDesc = new Filter("HSN_Description", FilterOperator.Contains, sValue);
      const oCombinedFilter = new Filter({
        filters: [oFilterCode, oFilterDesc],
        and: false
      });
      oBinding.filter([oCombinedFilter]);
    },
    onHSNValueHelpConfirm: function _onHSNValueHelpConfirm(oEvent) {
      const oSelectedItem = oEvent.getParameter("selectedItem");
      const oInput = this.getView()?.data("valueHelpSourceInput");
      if (!oSelectedItem) return;
      const oContext = oSelectedItem.getBindingContext("hsnModel");
      if (oContext) {
        const sSelectedCode = oContext.getProperty("HSN_CD");
        oInput.setValue(sSelectedCode);
      }
    },
    onShareToWhatsApp: async function _onShareToWhatsApp() {
      const oView = this.getView();
      if (!oView) return;
      const sSelectMode = oView.byId("mySelect").getSelectedItem()?.getText();
      const oModel = oView.getModel();
      const jspdfLib = window.jspdf;
      if (!jspdfLib) {
        MessageBox.error("jsPDF library is not loaded.");
        return;
      }
      let sClientName = "";
      let sTotalAmount = "";
      let sDocIdentifier = "";
      let sRawAddressText = "";
      let sTargetFilename = "Document.pdf";
      const doc = new jspdfLib.jsPDF();
      const pageWidth = doc.internal.pageSize.width;
      const pageHeight = doc.internal.pageSize.height;
      const margin = 5;
      doc.setDrawColor(0, 0, 0);
      doc.setLineWidth(0.3);
      doc.rect(margin, margin, pageWidth - margin * 2, pageHeight - margin * 2);
      if (this.sLogoBase64) {
        doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25);
      }
      doc.setFontSize(10);
      doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, {
        align: 'right'
      });
      doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, {
        align: 'right'
      });
      doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, {
        align: 'right'
      });
      doc.setFontSize(9);
      doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, {
        align: 'right'
      });
      doc.text("Nagarabhavi, Bangalore-560072", pageWidth - 14, 35, {
        align: 'right'
      });
      doc.line(5, 40, pageWidth - 5, 40);
      if (sSelectMode === "Quotation") {
        const oHeader = oModel.getProperty("/header");
        const aItems = oModel.getProperty("/products");
        sRawAddressText = oHeader.To || "";
        sClientName = this._getFirstLineName(sRawAddressText);
        sTotalAmount = formatINR(oModel.getProperty("/totalSum"));
        sDocIdentifier = "Quotation";
        sTargetFilename = `${sClientName}_Quotation.pdf`;
        doc.setFont("helvetica", "bold");
        doc.text(`Date: ${oHeader.Date}`, pageWidth - 14, 45, {
          align: 'right'
        });
        doc.text("To,", 14, 45);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.To, 80), 14, 50);
        doc.setFont("helvetica", "bold");
        let finalHeaderY = 70;
        doc.text("Sub: " + oHeader.Subject, 14, 70);
        doc.setFont("helvetica", "normal");
        if (oHeader.AddtionalInfo !== "") {
          doc.text(oHeader.AddtionalInfo, 14, finalHeaderY + 4);
        }
        const bShowGST = oView.byId("chkGST").getSelected();
        const bShowTotal = oView.byId("chkTotal").getSelected();
        const tableBody = aItems.map((item, index) => [index + 1, item.productName, item.quantity, item.price + item.symbol, formatINR(item.total)]);
        const subtotal = aItems.reduce((acc, cur) => acc + parseFloat(cur.total || 0), 0);
        const gstAmount = subtotal * 0.18;
        const grandTotal = subtotal + gstAmount;
        if (bShowGST) tableBody.push([{
          content: '18% GST Amount',
          colSpan: 4,
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }, {
          content: formatINR(gstAmount),
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }]);
        if (bShowTotal) tableBody.push([{
          content: 'Total',
          colSpan: 4,
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }, {
          content: formatINR(grandTotal),
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }]);
        doc.autoTable({
          startY: finalHeaderY + 8,
          head: [['Sl.No.', 'Particulars', 'Quantity', 'Rate', 'Total (Rs.)']],
          body: tableBody,
          theme: 'grid',
          styles: {
            fontSize: 9,
            lineColor: [0, 0, 0],
            lineWidth: 0.1
          },
          headStyles: {
            fillColor: [255, 255, 255],
            textColor: [0, 0, 0],
            fontStyle: 'bold'
          }
        });
        let finalY = doc.lastAutoTable.finalY;
        if (oHeader.TermsAndConditions !== "") {
          finalY += 10;
          doc.text(doc.splitTextToSize(oHeader.TermsAndConditions, pageWidth - 28), 14, finalY + 4);
        }
        if (oView.byId("chkBankDetail").getSelected()) {
          finalY += 20;
          doc.text(doc.splitTextToSize(oHeader.BankDetails, pageWidth - 28), 14, finalY + 4);
        }
      } else if (sSelectMode === "TAX-INVOICE") {
        const oData = oModel.getData();
        sRawAddressText = oData.taxHeader.To || "";
        sClientName = this._getFirstLineName(sRawAddressText);
        sTotalAmount = formatINR(oModel.getProperty("/taxtotalSum"));
        sDocIdentifier = `Tax Invoice (${oData.taxHeader.InvoiceNo || "N/A"})`;
        sTargetFilename = `${sClientName}_TaxInvoice.pdf`;
        doc.setFont("helvetica", "bold");
        doc.text("To,", 15, 44);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oData.taxHeader.To, 90), 15, 49);
        const metaX = 130;
        doc.text(`GST No: ${oData.taxHeader.GSTNo}`, metaX, 44);
        doc.text(`Invoice No: ${oData.taxHeader.InvoiceNo}`, metaX, 50);
        doc.text(`Date: ${oData.taxHeader.Date}`, metaX, 56);
        const tableRows = oData.taxProducts.map((item, index) => [index + 1, item.taxpProductName, item.taxHSNCode, formatINR(item.taxPrice) + item.taxSymbol, item.taxQuantity, formatINR(item.taxTotal)]);
        const totalAmount = oData.taxProducts.reduce((sum, item) => sum + parseFloat(item.taxTotal), 0);
        const taxVal = totalAmount * 0.09;
        const grandTotal = totalAmount + taxVal * 2;
        tableRows.push([{
          content: `TOTAL:`,
          colSpan: 5,
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }, {
          content: formatINR(totalAmount),
          styles: {
            halign: 'right',
            fontStyle: 'bold'
          }
        }], [{
          content: `GRAND TOTAL:`,
          colSpan: 5,
          styles: {
            halign: 'right',
            fontStyle: 'bold',
            fontSize: 10
          }
        }, {
          content: formatINR(grandTotal),
          styles: {
            halign: 'right',
            fontStyle: 'bold',
            fontSize: 10
          }
        }]);
        doc.autoTable({
          startY: 75,
          head: [["SI No.", "Particulars", "HSN Code", "Rate", "No.of Units", "Amount"]],
          body: tableRows,
          theme: 'grid'
        });
        const finalY = doc.lastAutoTable.finalY + 5;
        doc.text(`Rupees in words: ${this.numberToWords(grandTotal)} Only`, 10, finalY + 10);
      } else {
        const oData = oModel.getData();
        sRawAddressText = oData.cashHeader.cashTo || "";
        sClientName = this._getFirstLineName(sRawAddressText);
        sTotalAmount = formatINR(oModel.getProperty("/cashTotalSum"));
        sDocIdentifier = "Cash Bill";
        sTargetFilename = `${sClientName}_CashBill.pdf`;
        doc.setFont("helvetica", "bold");
        doc.text("To,", 15, 46);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oData.cashHeader.cashTo, 90), 15, 51);
        doc.text(`Date: ${oData.cashHeader.cashDate}`, pageWidth - 14, 46, {
          align: 'right'
        });
        const tableRows = oData.cashProducts.map((item, index) => [index + 1, item.cashBody, item.cashQuantity, formatINR(item.cashAmount)]);
        const totalAmount = oData.cashProducts.reduce((sum, item) => sum + parseFloat(item.cashAmount || 0), 0);
        tableRows.push([{
          content: `GRAND TOTAL:`,
          colSpan: 3,
          styles: {
            halign: 'right',
            fontStyle: 'bold',
            fontSize: 10
          }
        }, {
          content: formatINR(totalAmount),
          styles: {
            halign: 'right',
            fontStyle: 'bold',
            fontSize: 10
          }
        }]);
        doc.autoTable({
          startY: 75,
          head: [["SI No.", "Particulars", "Quantity", "Amount"]],
          body: tableRows,
          theme: 'grid'
        });
      }
      if (this.sSignaturBase64) {
        doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, pageHeight - 40, 70, 25);
      }
      const blobHTML5 = doc.output('blob');
      const pdfFile = new File([blobHTML5], sTargetFilename, {
        type: 'application/pdf'
      });
      if (navigator.canShare && navigator.canShare({
        files: [pdfFile]
      })) {
        try {
          await navigator.share({
            files: [pdfFile],
            title: `${sDocIdentifier} - IN-TELECOM SERVICES`,
            text: `Hello ${sClientName},\n\nYour *${sDocIdentifier}* from *IN-TELECOM SERVICES* is attached.\n\n*Amount Due:* Rs. ${sTotalAmount}\n\nThank you!`
          });
          MessageToast.show("Share prompt initialized successfully.");
        } catch (error) {
          MessageBox.error("Sharing processing was cancelled or encountered an error.");
        }
      } else {
        MessageBox.information("Direct file attachment is not supported on this browser window configuration. The PDF file will be downloaded onto your local system shelf. You can drag and drop it into your WhatsApp target chat window manually.", {
          title: "Browser Sharing Limitation",
          onClose: () => {
            doc.save(sTargetFilename);
            const aPhoneMatch = sRawAddressText.match(/(?:(?:\+91)|91)?\s*([6-9]\d{9})\b/);
            let sTargetPhone = aPhoneMatch ? "91" + aPhoneMatch[1] : "";
            const sMessageText = `Hello ${sClientName},\n\nYour *${sDocIdentifier}* from *IN-TELECOM SERVICES* is ready.\n\n*Amount Due:* Rs. ${sTotalAmount}`;
            sap.m.URLHelper.redirect(`https://api.whatsapp.com/send?phone=${sTargetPhone}&text=${encodeURIComponent(sMessageText)}`, true);
          }
        });
      }
    },
    onGenerateWord: async function _onGenerateWord(sMode) {
      if (!window.docxtemplater) {
        try {
          await new Promise((resolve, reject) => {
            const scriptZip = document.createElement("script");
            scriptZip.src = "https://cdn.jsdelivr.net/npm/pizzip@3.1.4/dist/pizzip.min.js";
            scriptZip.onload = () => {
              const scriptDocx = document.createElement("script");
              scriptDocx.src = "https://cdn.jsdelivr.net/npm/docxtemplater@3.45.0/build/docxtemplater.js";
              scriptDocx.onload = () => resolve();
              scriptDocx.onerror = err => reject(err);
              document.head.appendChild(scriptDocx);
            };
            scriptZip.onerror = err => reject(err);
            document.head.appendChild(scriptZip);
          });
        } catch (error) {
          MessageBox.error("Failed to dynamically load client-side template engines.");
          return;
        }
      }
      const PizZip = window.PizZip;
      const Docxtemplater = window.docxtemplater;
      const oModel = this.getView()?.getModel();
      const oData = oModel.getData();
      let sTemplatePath = "";
      let sTargetFilename = "";
      let oTemplateDataMap = {};
      if (sMode === "Quotation") {
        sTemplatePath = "my/app/generatebill/model/Quotation_Template.docx";
        sTargetFilename = this._getFirstLineName(oData.header.To) + "_Quotation.docx";
        const subtotal = oData.products.reduce((acc, cur) => acc + parseFloat(cur.total || 0), 0);
        const gstAmount = subtotal * 0.18;
        oTemplateDataMap = {
          To: oData.header.To,
          Date: oData.header.Date,
          Subject: oData.header.Subject,
          AdditionalInfo: oData.header.AddtionalInfo,
          prds: oData.products.map((item, idx) => ({
            i: idx + 1,
            productName: item.productName,
            quantity: item.quantity,
            price: item.price,
            symbol: item.symbol,
            total: formatINR(item.total)
          })),
          gstAmount: formatINR(gstAmount),
          grandTotal: formatINR(oModel.getProperty("/totalSum")),
          Notes: oData.header.Notes,
          TermsAndConditions: oData.header.TermsAndConditions
        };
      } else if (sMode === "TAX-INVOICE") {
        sTemplatePath = "my/app/generatebill/model/TaxInvoice_Template.docx";
        sTargetFilename = this._getFirstLineName(oData.taxHeader.To) + "_TaxInvoice.docx";
        const totalAmount = oData.taxProducts.reduce((sum, item) => sum + parseFloat(item.taxTotal), 0);
        const taxVal = totalAmount * 0.09;
        const grandTotal = totalAmount + taxVal * 2;
        oTemplateDataMap = {
          To: oData.taxHeader.To,
          InvoiceNo: oData.taxHeader.InvoiceNo,
          Date: oData.taxHeader.Date,
          PONo: oData.taxHeader.PONo,
          PODate: oData.taxHeader.PODate,
          taxPrd: oData.taxProducts.map((item, idx) => ({
            in: idx + 1,
            taxPart: item.taxpProductName,
            taxHSN: item.taxHSNCode,
            taxP: formatINR(item.taxPrice),
            taxS: item.taxSymbol,
            taxQ: item.taxQuantity,
            taxT: formatINR(item.taxTotal)
          })),
          total: formatINR(totalAmount),
          cgst: formatINR(taxVal),
          sgst: formatINR(taxVal),
          grandTotal: formatINR(grandTotal),
          words: this.numberToWords(grandTotal),
          PartyGST: oData.taxHeader.PartyGST
        };
      } else {
        sTemplatePath = "my/app/generatebill/model/CashBill_Template.docx";
        sTargetFilename = this._getFirstLineName(oData.cashHeader.cashTo) + "_CashBill.docx";
        const totalAmount = oData.cashProducts.reduce((sum, item) => sum + parseFloat(item.cashAmount || 0), 0);
        oTemplateDataMap = {
          cashTo: oData.cashHeader.cashTo,
          cashDate: oData.cashHeader.cashDate,
          cP: oData.cashProducts.map((item, idx) => ({
            i: idx + 1,
            cashBody: item.cashBody,
            cashQuantity: item.cashQuantity,
            cashAmount: formatINR(item.cashAmount)
          })),
          cashTotalSum: formatINR(totalAmount),
          words: this.numberToWords(totalAmount)
        };
      }
      try {
        const sTemplateUrl = sap.ui.require.toUrl(sTemplatePath);
        const response = await fetch(sTemplateUrl);
        if (!response.ok) throw new Error("Could not fetch the specified Word template.");
        const arrayBuffer = await response.arrayBuffer();
        const zip = new PizZip(arrayBuffer);
        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true
        });
        doc.setData(oTemplateDataMap);
        doc.render();
        const outBlob = doc.getZip().generate({
          type: "blob",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(outBlob);
        link.download = sTargetFilename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        MessageToast.show("Word document successfully downloaded from template mapping!");
      } catch (error) {
        MessageBox.error("Error filling out document template: " + error.message);
      }
    },
    /**
    * Asynchronously captures basic billing metrics and posts them to the dedicated analytics bin
    * @param {string} sDocType The category of document ("Quotation" | "TAX-INVOICE" | "Cash Bill")
    * @param {string} sDocNo Generated identification string
    * @param {string} sClient Target client name
    * @param {number} fAmount Grand total invoice revenue
    */
    _logDocumentAnalytics: async function _logDocumentAnalytics(sDocType, sDocNo, sClient, fAmount) {
      try {
        // 1. Fetch the existing collection matrix from the cloud bin
        const oGetResponse = await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}/latest`, {
          method: "GET",
          headers: {
            "X-Master-Key": this.sMasterKey
          }
        });
        if (!oGetResponse.ok) throw new Error("Failed to read analytical historical logs.");
        const oGetResult = await oGetResponse.json();
        const aHistory = oGetResult.record.records || [];

        // 2. Append the new structured transaction item details
        // Inside your _logDocumentAnalytics method, enhance the object pushed to aHistory:
        const oDate = new Date();
        const aMonths = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        aHistory.push({
          id: sDocNo,
          type: sDocType,
          client: sClient,
          amount: fAmount,
          timestamp: oDate.toISOString(),
          month: aMonths[oDate.getMonth()],
          // Stores e.g., "May"
          year: oDate.getFullYear().toString() // Stores e.g., "2026"
        });

        // 3. Commit the updated history collection block back to your cloud data container
        await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}`, {
          method: "PUT",
          headers: {
            "Content-Type": "application/json",
            "X-Master-Key": this.sMasterKey
          },
          body: JSON.stringify({
            records: aHistory
          })
        });
      } catch (oError) {
        console.error("Background analytical tracking logging engine failed: ", oError);
      }
    },
    /**
     * Fetches analytical transaction logs from your cloud database bin and presents
     * an interactive, filterable analytics dashboard with rolling dynamic year filters.
     */
    onShowAnalyticsGraph: async function _onShowAnalyticsGraph() {
      sap.ui.core.BusyIndicator.show(0);
      try {
        // 1. Fetch historical record arrays from your separate analytics bin container
        const response = await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}/latest`, {
          method: "GET",
          headers: {
            "X-Master-Key": this.sMasterKey
          }
        });
        if (!response.ok) throw new Error("Could not retrieve analytical transaction entries.");
        const result = await response.json();
        const aRecords = result.record.records || [];

        // 2. Initialize the dynamic UI container wrappers
        const oHtmlMetricsDashboard = new sap.ui.core.HTML();

        // 3. Calculate rolling calendar year strings dynamically (Current Year, -1, -2)
        const iCurrentYear = new Date().getFullYear();
        const sYearCurrent = iCurrentYear.toString();
        const sYearMinus1 = (iCurrentYear - 1).toString();
        const sYearMinus2 = (iCurrentYear - 2).toString();

        // 4. Define responsive dropdown selectors for Year and Month filtering
        const oYearSelect = new sap.m.Select({
          selectedKey: sYearCurrent,
          // Defaults to current year string
          items: [new sap.ui.core.Item({
            key: "ALL",
            text: "All Years"
          }), new sap.ui.core.Item({
            key: sYearMinus2,
            text: sYearMinus2
          }), new sap.ui.core.Item({
            key: sYearMinus1,
            text: sYearMinus1
          }), new sap.ui.core.Item({
            key: sYearCurrent,
            text: `${sYearCurrent} (Current Year)`
          })]
        });
        const oMonthSelect = new sap.m.Select({
          selectedKey: "ALL",
          // Defaults to all months string
          items: [new sap.ui.core.Item({
            key: "ALL",
            text: "All Months"
          }), ...["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"].map(sM => new sap.ui.core.Item({
            key: sM,
            text: sM
          }))]
        });
        const oFilterToolbar = new sap.m.HBox({
          width: "100%",
          justifyContent: "SpaceBetween",
          items: [oYearSelect, oMonthSelect]
        }).addStyleClass("sapUiSmallMarginBottom");

        // 5. Core calculation engine loop to process dropdown selections reactively
        const fnFilterAndRefreshDashboard = () => {
          const sSelYear = oYearSelect.getSelectedKey();
          const sSelMonth = oMonthSelect.getSelectedKey();

          // Filter dataset entries matching selected filter keys
          const aFiltered = aRecords.filter(rec => {
            const oDate = new Date(rec.timestamp);
            const aMonths = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            const sRowYear = rec.year || oDate.getFullYear().toString();
            const sRowMonth = rec.month || aMonths[oDate.getMonth()];
            const bYearMatch = sSelYear === "ALL" || sRowYear === sSelYear;
            const bMonthMatch = sSelMonth === "ALL" || sRowMonth === sSelMonth;
            return bYearMatch && bMonthMatch;
          });

          // Calculate aggregated numeric totals
          let fQuoteTotal = 0,
            fTaxTotal = 0,
            fCashTotal = 0;
          aFiltered.forEach(rec => {
            if (rec.type === "Quotation") fQuoteTotal += rec.amount;else if (rec.type === "TAX-INVOICE") fTaxTotal += rec.amount;else if (rec.type === "Cash Bill") fCashTotal += rec.amount;
          });
          const fOverallRevenue = fQuoteTotal + fTaxTotal + fCashTotal;

          // Safe percentage string generation logic
          const getPercentageString = fValue => {
            if (fOverallRevenue === 0) return "0%";
            return `${Math.min(100, Math.round(fValue / fOverallRevenue * 100))}%`;
          };
          const sQuoteWidth = getPercentageString(fQuoteTotal);
          const sTaxWidth = getPercentageString(fTaxTotal);
          const sCashWidth = getPercentageString(fCashTotal);

          // Dynamic HTML template compilation injection layer
          oHtmlMetricsDashboard.setContent(`
                <div style="font-family: Arial, sans-serif; padding: 5px; min-width: 360px; color: #333;">
                    <div style="margin-bottom: 8px; font-size: 13px; font-weight: bold; color: #555;">
                        Filtered Operations count: <span style="color: #2b7d2b;">${aFiltered.length} Documents</span>
                    </div>
                    <div style="margin-bottom: 18px; font-size: 15px; font-weight: bold; color: #000; border-bottom: 2px solid #eee; padding-bottom: 6px;">
                        Gross Segment Volume: <span style="color: #0a6ed1;">Rs. ${formatINR(fOverallRevenue)}</span>
                    </div>
                    
                    <div style="margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; font-size: 11px; margin-bottom: 4px; font-weight: bold;">
                            <span>QUOTATIONS (Rs. ${formatINR(fQuoteTotal)})</span>
                            <span>${sQuoteWidth}</span>
                        </div>
                        <div style="width: 100%; background: #e0e0e0; border-radius: 4px; height: 14px; overflow: hidden;">
                            <div style="width: ${sQuoteWidth}; background: #2b7d2b; height: 100%; border-radius: 4px 0 0 4px; transition: width 0.3s ease;"></div>
                        </div>
                    </div>

                    <div style="margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; font-size: 11px; margin-bottom: 4px; font-weight: bold;">
                            <span>TAX INVOICES (Rs. ${formatINR(fTaxTotal)})</span>
                            <span>${sTaxWidth}</span>
                        </div>
                        <div style="width: 100%; background: #e0e0e0; border-radius: 4px; height: 14px; overflow: hidden;">
                            <div style="width: ${sTaxWidth}; background: #e67e22; height: 100%; border-radius: 4px 0 0 4px; transition: width 0.3s ease;"></div>
                        </div>
                    </div>

                    <div style="margin-bottom: 4px;">
                        <div style="display: flex; justify-content: space-between; font-size: 11px; margin-bottom: 4px; font-weight: bold;">
                            <span>CASH BILLS (Rs. ${formatINR(fCashTotal)})</span>
                            <span>${sCashWidth}</span>
                        </div>
                        <div style="width: 100%; background: #e0e0e0; border-radius: 4px; height: 14px; overflow: hidden;">
                            <div style="width: ${sCashWidth}; background: #d32f2f; height: 100%; border-radius: 4px 0 0 4px; transition: width 0.3s ease;"></div>
                        </div>
                    </div>
                </div>
            `);
        };

        // CRITICAL CHANGE: Attach listeners AFTER the filter function is completely declared
        oYearSelect.attachChange(() => {
          fnFilterAndRefreshDashboard();
        });
        oMonthSelect.attachChange(() => {
          fnFilterAndRefreshDashboard();
        });

        // Fire initial load execution map to paint statistics immediately upon load
        fnFilterAndRefreshDashboard();

        // 6. Construct Primary Dialog Wrapper Overlay setup
        const oDialog = new Dialog({
          title: "Business Revenue Analytics Dashboard",
          type: "Standard",
          contentWidth: "420px",
          customHeader: new sap.m.Bar({
            contentLeft: [new sap.m.Title({
              text: "Business Revenue Analytics Dashboard"
            })],
            contentRight: [new sap.m.Button({
              icon: "sap-icon://decline",
              type: "Transparent",
              press: () => {
                oDialog.close();
              }
            })]
          }),
          content: [new sap.m.VBox({
            items: [oFilterToolbar, oHtmlMetricsDashboard]
          }).addStyleClass("sapUiContentPadding")],
          buttons: [new sap.m.Button({
            text: "Clear History",
            icon: "sap-icon://delete",
            type: "Negative",
            press: () => {
              this._clearAnalyticsCloudData(oDialog);
            }
          }), new sap.m.Button({
            text: "Download Excel",
            icon: "sap-icon://excel-attachment",
            type: "Accept",
            press: () => {
              this._downloadAnalyticsExcel(aRecords);
            }
          })],
          afterClose: () => {
            oDialog.destroy();
          }
        });
        oDialog.open();
      } catch (err) {
        sap.m.MessageBox.error(`Analytics Dashboard Generation Fault: ${err.message || err}`);
      } finally {
        sap.ui.core.BusyIndicator.hide();
      }
    },
    /**
    * Converts the current cloud analytics logs into an Excel file and downloads it.
    * @param {any[]} aRecords The raw array of billing logs fetched from the bin
    */
    _downloadAnalyticsExcel: function _downloadAnalyticsExcel(aRecords) {
      if (!aRecords || aRecords.length === 0) {
        sap.m.MessageToast.show("No transaction records available to export.");
        return;
      }

      // 1. Map the clean cloud fields to descriptive Excel columns
      const aExcelRows = aRecords.map(rec => ({
        "Document ID": rec.id,
        "Document Type": rec.type,
        "Client / Customer Name": rec.client,
        "Grand Total (Rs.)": rec.amount,
        "Generated Timestamp": rec.timestamp
      }));

      // 2. Build and export the spreadsheet using your existing XLSX library reference
      const ws = XLSX.utils.json_to_sheet(aExcelRows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Revenue Overview");
      XLSX.writeFile(wb, "Business_Analytics_Summary.xlsx");
      sap.m.MessageToast.show("Analytics data exported to Excel successfully!");
    },
    // /**
    //  * Wipes out all historical transactions inside the dedicated analytics cloud bin.
    /**
    * Prompts a password validation dialog box before wiping historical records from the cloud bin.
    * @param {sap.m.Dialog} oDashboardDialog Reference to the parent dashboard dialog overlay
    */
    _clearAnalyticsCloudData: function _clearAnalyticsCloudData(oDashboardDialog) {
      // 1. Hardcode your chosen administrative access password here
      const sMasterPassword = "clearMe@22";

      // 2. Instantiate a secure Password Input field component
      const oPasswordInput = new sap.m.Input({
        type: sap.m.InputType.Password,
        placeholder: "Enter admin password",
        width: "100%"
      });

      // 3. Build the primary security authentication container window
      const oSecurityDialog = new Dialog({
        title: "Security Verification Required",
        type: "Message",
        state: "Warning",
        // Provides a distinct orange accent bar
        content: [new sap.m.Text({
          text: "This action will permanently wipe out all recorded historical transaction data. Please enter the master password to authorize this action:"
        }), oPasswordInput],
        beginButton: new sap.m.Button({
          text: "Authorize & Clear",
          type: "Reject",
          // Highlights the button red for destructive operations
          press: async () => {
            const sEnteredPassword = oPasswordInput.getValue();

            // 4. Validate input value against the hardcoded key string
            if (sEnteredPassword !== sMasterPassword) {
              sap.m.MessageBox.error("Authentication Failed! Incorrect password entry.");
              oPasswordInput.setValue(""); // Reset input field clear
              return;
            }

            // 5. If correct, proceed to trigger your existing database reset logic
            oSecurityDialog.close();
            sap.ui.core.BusyIndicator.show(0);
            try {
              const response = await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}`, {
                method: "PUT",
                headers: {
                  "Content-Type": "application/json",
                  "X-Master-Key": this.sMasterKey
                },
                body: JSON.stringify({
                  records: []
                })
              });
              if (!response.ok) throw new Error("Cloud database rejected the wipe request.");
              sap.m.MessageToast.show("Analytics history successfully cleared from the cloud!");
              oDashboardDialog.close(); // Closes the underlying dashboard chart layout cleanly
            } catch (err) {
              sap.m.MessageBox.error(`Clear Operational Fault: ${err.message}`);
            } finally {
              sap.ui.core.BusyIndicator.hide();
            }
          }
        }),
        endButton: new sap.m.Button({
          text: "Abort",
          press: () => {
            oSecurityDialog.close();
          }
        }),
        afterClose: () => {
          oSecurityDialog.destroy();
        }
      });
      oSecurityDialog.open();
    }
  });
  return View;
});
//# sourceMappingURL=View-dbg.controller.js.map
