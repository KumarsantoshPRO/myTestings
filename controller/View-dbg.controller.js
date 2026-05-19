sap.ui.define(["sap/ui/core/mvc/Controller", "sap/ui/model/json/JSONModel", "sap/m/MessageBox", "sap/m/MessageToast"], function (Controller, JSONModel, MessageBox, MessageToast) {
  "use strict";

  // External SheetJS Library Reference

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
      this.sLogoBase64 = "";
      this.sSignaturBase64 = "";
      this.sStorageKey = "my_billing_app_draft_data";
      this.sSequenceKey = "my_billing_app_invoice_seq_counter";
    },
    onInit: function _onInit() {
      let oData;
      const sSavedData = localStorage.getItem(this.sStorageKey);
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
            Date: "",
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
            Date: "",
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
            cashDate: "",
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

      // Setup original visibility views state 
      this._updateSectionVisibilities("Quotation");
    },
    _getFirstLineName: function _getFirstLineName(sAddress) {
      if (!sAddress || sAddress.trim() === "") {
        return "Document";
      }
      return sAddress.split("\n")[0].replace(/[/\\?%*:|"<>\s]/g, "_").trim();
    },
    onGenerateNextInvoiceNumber: function _onGenerateNextInvoiceNumber() {
      const oModel = this.getView()?.getModel();

      // 1. Determine Current Indian Financial Year (Changes on April 1st)
      const oToday = new Date();
      const iCurrentMonth = oToday.getMonth(); // 0 = January, 3 = April
      const iCurrentYear = oToday.getFullYear();
      let iStartFY;
      let iEndFY;
      if (iCurrentMonth >= 3) {
        // April to December
        iStartFY = iCurrentYear;
        iEndFY = iCurrentYear + 1;
      } else {
        // January to March
        iStartFY = iCurrentYear - 1;
        iEndFY = iCurrentYear;
      }

      // Format years into "26-27" format
      const sStartFYString = iStartFY.toString().substring(2);
      const sEndFYString = iEndFY.toString().substring(2);
      const sFinancialYearLabel = `${sStartFYString}-${sEndFYString}`; // e.g., "26-27"

      // 2. Setup a unique sequence counter storage key for this specific FY
      const sFYSpecificStorageKey = `my_billing_app_sequence_${sFinancialYearLabel}`;

      // 3. Get current sequence value or default to 0 if it's a new Financial Year
      let iLastSequence = parseInt(localStorage.getItem(sFYSpecificStorageKey) || "0");
      let iNextSequence = iLastSequence + 1;

      // Save updated sequence back to storage
      localStorage.setItem(sFYSpecificStorageKey, iNextSequence.toString());

      // 4. Pad single-digit invoice sequences with a leading zero (e.g., 1 -> "01", 10 -> "10")
      const sPaddedSequence = iNextSequence < 10 ? `0${iNextSequence}` : `${iNextSequence}`;

      // 5. Construct final string: e.g., "09/26-27" or "100/26-27"
      const sNewInvoiceNo = `${sPaddedSequence}/${sFinancialYearLabel}`;

      // Update UI layout bindings
      oModel.setProperty("/taxHeader/InvoiceNo", sNewInvoiceNo);
      MessageToast.show(`Sequence generated for FY ${sFinancialYearLabel}!`);
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
    },
    // --- Core Dynamic Excel Processing Handling ---
    onExcelDownload: function _onExcelDownload(oEvent) {
      const sSelectMode = (this.getView()?.byId("mySelect")).getSelectedItem()?.getText();
      const oModel = this.getView()?.getModel();
      let dataToExport = [];
      let sFileName = "Export.xlsx";
      if (sSelectMode === "Quotation") {
        sFileName = this._getFirstLineName(oModel.getProperty("/header/To")) + "_QuotationItems.xlsx";
        dataToExport = oModel.getProperty("/products").map(item => ({
          "Particulars": item.productName,
          "Rate": item.price,
          "Quantity": item.quantity,
          "Symbol": item.symbol,
          "Total": item.total
        }));
      } else if (sSelectMode === "TAX-INVOICE") {
        sFileName = this._getFirstLineName(oModel.getProperty("/taxHeader/To")) + "_TaxInvoiceItems.xlsx";
        dataToExport = oModel.getProperty("/taxProducts").map(item => ({
          "Particulars": item.taxpProductName,
          "HSN Code": item.taxHSNCode,
          "Rate": item.taxPrice,
          "Quantity": item.taxQuantity,
          "Symbol": item.taxSymbol,
          "Total": item.taxTotal
        }));
      } else {
        sFileName = this._getFirstLineName(oModel.getProperty("/cashHeader/cashTo")) + "_CashBillItems.xlsx";
        dataToExport = oModel.getProperty("/cashProducts").map(item => ({
          "Particulars": item.cashBody,
          "Quantity": item.cashQuantity,
          "Amount": item.cashAmount
        }));
      }
      const ws = XLSX.utils.json_to_sheet(dataToExport);
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
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        const sSelectMode = (this.getView()?.byId("mySelect")).getSelectedItem()?.getText();
        const oModel = this.getView()?.getModel();
        if (sSelectMode === "Quotation") {
          const parsedProducts = jsonData.map(row => ({
            productName: row["Particulars"] || "",
            price: row["Rate"]?.toString() || "0.00",
            quantity: parseInt(row["Quantity"]) || 1,
            symbol: row["Symbol"] || "",
            total: row["Total"]?.toString() || "0.00"
          }));
          oModel.setProperty("/products", parsedProducts);
          this.onCalc();
        } else if (sSelectMode === "TAX-INVOICE") {
          const parsedTax = jsonData.map(row => ({
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
          const parsedCash = jsonData.map(row => ({
            cashBody: row["Particulars"] || "",
            cashQuantity: row["Quantity"]?.toString() || "1",
            cashAmount: row["Amount"]?.toString() || "0.00"
          }));
          oModel.setProperty("/cashProducts", parsedCash);
          this.onCashCalc();
        }
        MessageToast.show("Excel rows processed successfully!");
        oFileUploader.clear();
      };
      reader.readAsArrayBuffer(oFile);
    },
    // --- MS Word native container engine logic handler implementation ---
    onExportToMSWord: function _onExportToMSWord() {
      const sSelectMode = (this.getView()?.byId("mySelect")).getSelectedItem()?.getText();
      const oModel = this.getView()?.getModel();
      let sHtmlContent = "";
      let sFileName = "Export.doc";

      // Common CSS Styles shared across all profiles to mimic the PDF design look
      const sStyleHeader = `
        <style>
            body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #000000; font-size: 10pt; line-height: 1.4; }
            .letterhead-table { width: 100%; border: none; margin-bottom: 20px; }
            .company-title { font-size: 18pt; font-weight: bold; color: #000000; font-family: Arial, sans-serif; }
            .company-details { font-size: 9.5pt; text-align: right; color: #333333; }
            .divider { border-top: 1.5pt solid #000000; margin-top: 10px; margin-bottom: 20px; }
            .doc-title { font-size: 14pt; font-weight: bold; text-align: center; margin-bottom: 20px; text-transform: uppercase; letter-spacing: 1px; }
            .metadata-table { width: 100%; border: none; margin-bottom: 25px; }
            .metadata-label { font-weight: bold; width: 15%; vertical-align: top; font-size: 10pt; }
            .metadata-value { width: 35%; vertical-align: top; font-size: 10pt; }
            .items-table { width: 100%; border-collapse: collapse; margin-bottom: 25px; }
            .items-table th { background-color: #FFFFFF; color: #000000; font-weight: bold; text-align: center; border: 0.5pt solid #000000; padding: 6px; font-size: 9.5pt; }
            .items-table td { border: 0.5pt solid #000000; padding: 6px; font-size: 9.5pt; vertical-align: top; }
            .text-center { text-align: center; }
            .text-right { text-align: right; }
            .total-row-label { font-weight: bold; text-align: right; border: 0.5pt solid #000000 !important; }
            .total-row-value { font-weight: bold; text-align: right; border: 0.5pt solid #000000 !important; }
            .section-heading { font-size: 10pt; font-weight: bold; margin-top: 20px; margin-bottom: 5px; text-decoration: underline; }
            .footer-notes { font-size: 9.5pt; color: #222222; margin-bottom: 15px; }
            .signature-space { margin-top: 50px; text-align: right; font-weight: bold; font-size: 10pt; }
        </style>
    `;

      // 1. Business Letterhead Layout Template Fragment (Matches your PDF Header exactly)
      const sLetterheadHtml = `
        <table class="letterhead-table" cellspacing="0" cellpadding="0">
            <tr>
                <td style="width: 50%; vertical-align: middle;">
                    <span class="company-title">IN-TELECOM SERVICES</span>
                </td>
                <td class="company-details" style="width: 50%; vertical-align: top;">
                    <b>Ph: (Off.):</b> 2348 1249<br>
                    97400 27266 / 98442 11193<br>
                    <b>E-mail:</b> intelecompatil@rediffmail.com<br>
                    #249, 7th Main, 4th Cross, 2nd Stage,<br>
                    Nagarabhavi, Bangalore-560072
                </td>
            </tr>
        </table>
        <div class="divider"></div>
    `;
      if (sSelectMode === "Quotation") {
        const oHeader = oModel.getProperty("/header");
        sFileName = this._getFirstLineName(oHeader.To) + "_Quotation.doc";
        sHtmlContent = `
            ${sStyleHeader}
            ${sLetterheadHtml}
            <div class="doc-title">QUOTATION</div>
            
            <table class="metadata-table" cellspacing="0" cellpadding="3">
                <tr>
                    <td class="metadata-label">To,</td>
                    <td class="metadata-value" rowspan="2">${oHeader.To.replace(/\n/g, '<br>')}</td>
                    <td class="metadata-label" style="text-align: right;">Date:</td>
                    <td class="metadata-value">${oHeader.Date}</td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                    <td colspan="2">&nbsp;</td>
                </tr>
            </table>

            <p style="font-size: 10pt; margin-bottom: 15px;"><b>Sub:</b> ${oHeader.Subject}</p>
            ${oHeader.AddtionalInfo ? `<p style="font-size: 10pt; margin-bottom: 20px; padding-left: 30px;">${oHeader.AddtionalInfo}</p>` : ''}

            <table class="items-table">
                <thead>
                    <tr>
                        <th style="width: 8%;">Sl.No.</th>
                        <th style="width: 52%;">Particulars</th>
                        <th style="width: 12%;">Quantity</th>
                        <th style="width: 13%;">Rate</th>
                        <th style="width: 15%;">Total (Rs.)</th>
                    </tr>
                </thead>
                <tbody>
                    ${oModel.getProperty("/products").map((item, idx) => `
                        <tr>
                            <td class="text-center">${idx + 1}</td>
                            <td>${item.productName.replace(/\n/g, '<br>')}</td>
                            <td class="text-center">${item.quantity}</td>
                            <td class="text-right">${item.price}${item.symbol}</td>
                            <td class="text-right">${formatINR(item.total)}</td>
                        </tr>
                    `).join('')}
                    
                    <tr>
                        <td colspan="4" class="total-row-label">18% GST Amount</td>
                        <td class="total-row-value">${formatINR(oModel.getProperty("/gst"))}</td>
                    </tr>
                    <tr>
                        <td colspan="4" class="total-row-label">Total</td>
                        <td class="total-row-value">${formatINR(oModel.getProperty("/totalSum"))}</td>
                    </tr>
                </tbody>
            </table>

            ${oHeader.TermsAndConditions ? `<div class="section-heading">Terms & Conditions:</div><div class="footer-notes">${oHeader.TermsAndConditions.replace(/\n/g, '<br>')}</div>` : ''}
            ${oHeader.Notes ? `<div class="section-heading">Notes:</div><div class="footer-notes">${oHeader.Notes.replace(/\n/g, '<br>')}</div>` : ''}
            
            <div class="section-heading">Bank Details:</div>
            <div class="footer-notes">${oHeader.BankDetails.replace(/\n/g, '<br>')}</div>

            <div class="signature-space">
                <p>For IN-TELECOM SERVICES</p>
                <br><br><br>
                <p>Authorized Signatory</p>
            </div>
        `;
      } else if (sSelectMode === "TAX-INVOICE") {
        const oHeader = oModel.getProperty("/taxHeader");
        sFileName = this._getFirstLineName(oHeader.To) + "_TaxInvoice.doc";
        sHtmlContent = `
            ${sStyleHeader}
            ${sLetterheadHtml}
            <div class="doc-title">TAX INVOICE</div>
            
            <table class="metadata-table" cellspacing="0" cellpadding="3">
                <tr>
                    <td class="metadata-label">To,</td>
                    <td class="metadata-value" style="width: 45%;" rowspan="5">${oHeader.To.replace(/\n/g, '<br>')}</td>
                    <td class="metadata-label" style="text-align: right; width: 20%;">GST No:</td>
                    <td class="metadata-value" style="width: 20%;">${oHeader.GSTNo}</td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                    <td class="metadata-label" style="text-align: right;">Invoice No:</td>
                    <td><b>${oHeader.InvoiceNo}</b></td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                    <td class="metadata-label" style="text-align: right;">Date:</td>
                    <td>${oHeader.Date}</td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                    <td class="metadata-label" style="text-align: right;">PO No:</td>
                    <td>${oHeader.PONo}</td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                    <td class="metadata-label" style="text-align: right;">PO Date:</td>
                    <td>${oHeader.PODate}</td>
                </tr>
            </table>

            <table class="items-table">
                <thead>
                    <tr>
                        <th style="width: 8%;">SI No.</th>
                        <th style="width: 42%;">Particulars</th>
                        <th style="width: 12%;">HSN Code</th>
                        <th style="width: 13%;">Rate</th>
                        <th style="width: 10%;">No.of Units</th>
                        <th style="width: 15%;">Amount</th>
                    </tr>
                </thead>
                <tbody>
                    ${oModel.getProperty("/taxProducts").map((item, idx) => `
                        <tr>
                            <td class="text-center">${idx + 1}</td>
                            <td>${item.taxpProductName.replace(/\n/g, '<br>')}</td>
                            <td class="text-center">${item.taxHSNCode}</td>
                            <td class="text-right">${formatINR(item.taxPrice)}${item.taxSymbol}</td>
                            <td class="text-center">${item.taxQuantity}</td>
                            <td class="text-right">${formatINR(item.taxTotal)}</td>
                        </tr>
                    `).join('')}
                    
                    <tr>
                        <td colspan="5" class="total-row-label">TOTAL:</td>
                        <td class="total-row-value">${formatINR(oModel.getProperty("/taxtotal"))}</td>
                    </tr>
                    <tr>
                        <td colspan="5" class="total-row-label">CGST @ 9%:</td>
                        <td class="total-row-value">${formatINR(oModel.getProperty("/taxcgst"))}</td>
                    </tr>
                    <tr>
                        <td colspan="5" class="total-row-label">SGST @ 9%:</td>
                        <td class="total-row-value">${formatINR(oModel.getProperty("/taxsgst"))}</td>
                    </tr>
                    <tr>
                        <td colspan="5" class="total-row-label" style="font-size:10.5pt;">GRAND TOTAL:</td>
                        <td class="total-row-value" style="font-size:10.5pt;">${formatINR(oModel.getProperty("/taxtotalSum"))}</td>
                    </tr>
                </tbody>
            </table>

            <p style="font-size: 10pt; margin-bottom: 5px;"><b>Rupees in words:</b> ${this.numberToWords(parseFloat(oModel.getProperty("/taxtotalSum")))} Only</p>
            <p style="font-size: 10pt; margin-bottom: 20px;"><b>Party GST No:</b> ${oHeader.PartyGST || ""}</p>

            <div class="section-heading">Bank Details:</div>
            <div class="footer-notes">${oHeader.BankDetails.replace(/\n/g, '<br>')}</div>

            <div class="signature-space">
                <p>For IN-TELECOM SERVICES</p>
                <br><br><br>
                <p>Authorized Signatory</p>
            </div>
        `;
      } else {
        const oHeader = oModel.getProperty("/cashHeader");
        sFileName = this._getFirstLineName(oHeader.cashTo) + "_CashBill.doc";
        sHtmlContent = `
            ${sStyleHeader}
            ${sLetterheadHtml}
            <div class="doc-title">CASH BILL</div>
            
            <table class="metadata-table" cellspacing="0" cellpadding="3">
                <tr>
                    <td class="metadata-label">To,</td>
                    <td class="metadata-value" rowspan="2">${oHeader.cashTo.replace(/\n/g, '<br>')}</td>
                    <td class="metadata-label" style="text-align: right;">Date:</td>
                    <td class="metadata-value">${oHeader.cashDate}</td>
                </tr>
                <tr>
                    <td>&nbsp;</td>
                    <td colspan="2">&nbsp;</td>
                </tr>
            </table>

            <table class="items-table">
                <thead>
                    <tr>
                        <th style="width: 10%;">SI No.</th>
                        <th style="width: 55%;">Particulars</th>
                        <th style="width: 15%;">Quantity</th>
                        <th style="width: 20%;">Amount</th>
                    </tr>
                </thead>
                <tbody>
                    ${oModel.getProperty("/cashProducts").map((item, idx) => `
                        <tr>
                            <td class="text-center">${idx + 1}</td>
                            <td>${item.cashBody.replace(/\n/g, '<br>')}</td>
                            <td class="text-center">${item.cashQuantity}</td>
                            <td class="text-right">${formatINR(item.cashAmount)}</td>
                        </tr>
                    `).join('')}
                    
                    <tr>
                        <td colspan="3" class="total-row-label" style="font-size:10.5pt;">GRAND TOTAL:</td>
                        <td class="total-row-value" style="font-size:10.5pt;">${formatINR(oModel.getProperty("/cashTotalSum"))}</td>
                    </tr>
                </tbody>
            </table>

            <p style="font-size: 10pt; margin-bottom: 20px;"><b>Rupees in words:</b> ${this.numberToWords(parseFloat(oModel.getProperty("/cashTotalSum")))} Only</p>

            <div class="section-heading">Bank Details:</div>
            <div class="footer-notes">${oHeader.cashbankDetails.replace(/\n/g, '<br>')}</div>

            <div class="signature-space">
                <p>For IN-TELECOM SERVICES</p>
                <br><br><br>
                <p>Authorized Signatory</p>
            </div>
        `;
      }

      // Wrap with the specific MS Word XML structure namespaces so Word handles the styling accurately
      const sFullBlobTemplate = `
        <html xmlns:o='urn:schemas-microsoft-com:office:office' 
              xmlns:w='urn:schemas-microsoft-com:office:word' 
              xmlns='http://www.w3.org/TR/REC-html40'>
        <head>
            </head>
        <body>
            <div style="margin: 0.5in 0.5in 0.5in 0.5in;">
                ${sHtmlContent}
            </div>
        </body>
        </html>
    `;
      const blob = new Blob([sFullBlobTemplate], {
        type: 'application/msword'
      });
      const URL = window.URL || window.webkitURL;
      const downloadLink = document.createElement("a");
      downloadLink.href = URL.createObjectURL(blob);
      downloadLink.download = sFileName;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      document.body.removeChild(downloadLink);
    },
    numberToWords: function _numberToWords(num) {
      const a = ['', 'One ', 'Two ', 'Three ', 'Four ', 'Five ', 'Six ', 'Seven ', 'Eight ', 'Nine ', 'Ten ', 'Eleven ', 'Twelve ', 'Thirteen ', 'Fourteen ', 'Fifteen ', 'Sixteen ', 'Seventeen ', 'Eighteen ', 'Nineteen '];
      const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
      if (num === 0) return 'Zero';
      const convertWholeNumber = amountStr => {
        const n = ('000000000' + amountStr).substr(-9).match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
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
    }
  });
  return View;
});
//# sourceMappingURL=View-dbg.controller.js.map
