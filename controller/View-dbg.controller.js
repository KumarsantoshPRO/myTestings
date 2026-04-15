sap.ui.define(["sap/ui/core/mvc/Controller", "sap/ui/model/json/JSONModel", "sap/m/MessageBox"], function (Controller, JSONModel, MessageBox) {
  "use strict";

  const View = Controller.extend("webapp.controller.View", {
    constructor: function constructor() {
      Controller.prototype.constructor.apply(this, arguments);
      this.sLogoBase64 = "";
      this.sSignaturBase64 = "";
    },
    onInit: function _onInit() {
      const oData = {
        header: {
          To: "",
          Date: "",
          Subject: "",
          Notes: "",
          TermsAndConditions: "",
          BankDetails: "Payment Mode: Via Online\nBank: State Bank of India,\nBranch: Mallathahalli Branch\nName: In - Telecom Services\nC/A No: 64064045533\nIFSC Code: SBIN0040457"
        },
        products: [{
          productName: "",
          quantity: 0,
          price: 0,
          total: "0.00"
        }],
        taxHeader: {
          To: "",
          GSTNo: "",
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
          taxQuantity: 0,
          taxPrice: 0,
          taxTotal: "0.00"
        }]
      };
      this.getView()?.setModel(new JSONModel(oData));
      this._loadLocalLogo("img/logo.jpg");
      this._loadSignature("img/Signature.jpg");
      // (this.getView()?.byId("idPage") as Page).setTitle("Quotation");
      (this.getView()?.byId("idBtnQuotation")).setVisible(true);
      (this.getView()?.byId("idBtnInvoice")).setVisible(false);
      (this.getView()?.byId("idOPSQuoteTaxInc")).setVisible(false);
      (this.getView()?.byId("idTaxSec")).setVisible(false);
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
    onAddRow: function _onAddRow() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/products");
      aProducts.push({
        productName: "",
        price: 0,
        quantity: 1,
        total: "0.00"
      });
      oModel.setProperty("/products", aProducts);
    },
    onDelete: function _onDelete(oEvent) {
      // 1. Get the item (row) that was clicked
      const oItemToDelete = oEvent.getParameter("listItem");

      // 2. Get the binding context path (e.g., "/products/2")
      const sPath = oItemToDelete.getBindingContext().getPath();

      // 3. Extract the index from the path
      const iIndex = parseInt(sPath.split("/").pop());

      // 4. Get the model and the data array
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/products");

      // 5. Remove the element at the specific index
      aProducts.splice(iIndex, 1);

      // 6. Update the model to refresh the UI
      oModel.setProperty("/products", aProducts);
      this.onCalc();
    },
    onCalc: function _onCalc() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/products");
      let fGrandTotal = 0;
      let gst = 0;
      let gstAmount = 0;
      aProducts.forEach(oProduct => {
        // 1. Calculate individual row total
        const fQty = oProduct.quantity;
        const fPrice = parseFloat(oProduct.price) || 0;
        let fRowTotal = 0;
        if (typeof fQty === "string") {
          fRowTotal = fPrice;
        } else if (typeof fQty === "number") {
          fRowTotal = fQty * fPrice;
        }

        // Update the row total in the model
        oProduct.total = fRowTotal.toFixed(2);

        // 2. Add to Grand Total
        gstAmount = gstAmount + fRowTotal * 0.18;
        fGrandTotal += fRowTotal;
      });

      // 3. Set the Grand Total property for the footer
      oModel.setProperty("/gst", gstAmount.toFixed(2));
      oModel.setProperty("/totalSum", (fGrandTotal + gstAmount).toFixed(2).toString());

      // Refresh model to ensure UI updates
      oModel.refresh();
    },
    validateForm: function _validateForm() {
      const oModel = this.getView()?.getModel();
      const oData = oModel.getData();

      // 1. Validate Header Fields
      const oHeader = oData.header;
      for (const key in oHeader) {
        if (key !== "Notes") {
          if (!oHeader[key] || oHeader[key].trim() === "") {
            MessageBox.error(`Please fill in the header field: ${key}`);
            return false;
          }
        }
      }

      // 2. Validate Products Array
      const aProducts = oData.products;
      if (aProducts.length === 0) {
        MessageBox.error("Please add at least one product.");
        return false;
      }
      for (let i = 0; i < aProducts.length; i++) {
        const oProd = aProducts[i];
        if (!oProd.productName || oProd.productName.trim() === "") {
          MessageBox.error(`Row ${i + 1}: Product Name is required.`);
          return false;
        }
        if (oProd.quantity <= 0) {
          MessageBox.error(`Row ${i + 1}: Quantity must be greater than 0.`);
          return false;
        }
      }
      return true; // All checks passed
    },
    onGeneratePDF: function _onGeneratePDF() {
      if (this.validateForm()) {
        const jspdfLib = window.jspdf;
        if (!jspdfLib) return;
        const oModel = this.getView()?.getModel();
        const oHeader = oModel.getProperty("/header");
        const aItems = oModel.getProperty("/products");
        const doc = new jspdfLib.jsPDF();
        const pageWidth = doc.internal.pageSize.width;
        const pageHeight = doc.internal.pageSize.height;
        const margin = 5;

        // --- 0. SET PAGE BORDER ---
        // rect(x, y, width, height)
        doc.setDrawColor(0, 0, 0); // Black border
        doc.setLineWidth(0.3);
        doc.rect(margin, margin, pageWidth - margin * 2, pageHeight - margin * 2);

        // HEADER SECTION
        if (this.sLogoBase64) {
          doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25);
        }
        doc.setTextColor(0, 0, 0);
        doc.setFontSize(9);

        // Company Contact Info (Right Aligned)
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

        // TO / SUB / DATE SECTION 
        doc.setFont("helvetica", "bold");
        doc.text(`Date: ${oHeader.Date}`, pageWidth - 14, 45, {
          align: 'right'
        });
        doc.text("To,", 14, 45);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.To, 80), 14, 50);
        doc.setFont("helvetica", "bold");
        doc.text("Sub: " + oHeader.Subject, 14, 70);

        //TABLE SECTION
        const bShowGST = (this.getView()?.byId("chkGST")).getSelected();
        const bShowTotal = (this.getView()?.byId("chkTotal")).getSelected();
        // 1. Prepare standard product rows
        const tableBody = aItems.map((item, index) => [index + 1, item.productName, item.quantity, Number(item.price).toFixed(2), Number(item.total).toFixed(2)]);

        // 2. Calculate Totals
        const subtotal = aItems.reduce((acc, cur) => acc + parseFloat(cur.total || 0), 0);
        const gstAmount = subtotal * 0.18;
        const grandTotal = subtotal + gstAmount;

        // 3. Add Total rows to the table array
        // We use colSpan: 4 to merge the first 4 columns into one single label cell
        if (bShowGST && bShowTotal) {
          tableBody.push([{
            content: '18% GST Amount',
            colSpan: 4,
            styles: {
              halign: 'right',
              fontStyle: 'bold'
            }
          }, {
            content: gstAmount.toFixed(2),
            styles: {
              halign: 'right',
              fontStyle: 'bold'
            }
          }], [{
            content: 'Total',
            colSpan: 4,
            styles: {
              halign: 'right',
              fontStyle: 'bold'
            }
          }, {
            content: grandTotal.toFixed(2),
            styles: {
              halign: 'right',
              fontStyle: 'bold'
            }
          }]);
        } else if (bShowGST) {
          tableBody.push([{
            content: '18% GST Amount',
            colSpan: 4,
            styles: {
              halign: 'right',
              fontStyle: 'bold'
            }
          }, {
            content: gstAmount.toFixed(2),
            styles: {
              halign: 'right',
              fontStyle: 'bold'
            }
          }]);
        } else if (bShowTotal) {
          tableBody.push([{
            content: 'Total',
            colSpan: 4,
            styles: {
              halign: 'right',
              fontStyle: 'bold'
            }
          }, {
            content: grandTotal.toFixed(2),
            styles: {
              halign: 'right',
              fontStyle: 'bold'
            }
          }]);
        }

        // 4. Generate the Table
        doc.autoTable({
          startY: 72,
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

        // Subtotal and GST Note
        // Get the checkbox states
        // const bShowGST = (this.getView()?.byId("chkGST") as CheckBox).getSelected();
        // const bShowTotal = (this.getView()?.byId("chkTotal") as CheckBox).getSelected();

        // doc.setFont("helvetica", "bold");
        // // Conditional Logic for GST Amount

        // if (bShowGST) {
        //     finalY = finalY + 4;
        //     doc.text("18% GST Amount", 14, finalY); //[span_16](end_span)
        //     doc.text(gstAmount.toFixed(2), pageWidth - 14, finalY, { align: 'right' });
        // }
        // // Conditional Logic for Grand Total
        // if (bShowTotal) {
        //     finalY = finalY + 4;
        //     doc.text("Total", 14, finalY); //[span_16](end_span)
        //     doc.text(grandTotal.toFixed(2), pageWidth - 14, finalY, { align: 'right' });
        // }

        // Terms & Conditions ---
        finalY = finalY + 10;
        doc.setFontSize(9);
        doc.setFont("helvetica", "bold");
        doc.text("Terms & Conditions:", 14, finalY);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.TermsAndConditions, pageWidth - 28), 14, finalY + 4);

        // Notes
        if (oHeader.Notes !== "") {
          finalY = finalY + 30;
          doc.setFont("helvetica", "bold");
          doc.text("Notes:", 14, finalY);
          doc.setFont("helvetica", "normal");
          doc.text(doc.splitTextToSize(oHeader.Notes, pageWidth - 28), 14, finalY + 4);
        }

        // Bank Details ---
        const bBankDetails = (this.getView()?.byId("chkBankDetail")).getSelected();
        if (bBankDetails) {
          finalY = finalY + 30;
          doc.setFontSize(9);
          doc.setFont("helvetica", "bold");
          doc.text("Bank Details:", 14, finalY);
          doc.setFont("helvetica", "normal");
          doc.text(doc.splitTextToSize(oHeader.BankDetails, pageWidth - 28), 14, finalY + 4);
        }

        // SIGNATURE SECTION
        const sigY = doc.internal.pageSize.height - 40;
        if (this.sSignaturBase64) {
          doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, sigY, 70, 25);
        }
        window.open(doc.output("bloburl"), "_blank");
      }
    },
    onRadioSelect: function _onRadioSelect(oEvent) {
      // Get the index of the selected button (0 for first, 1 for second)
      const iSelectedIndex = oEvent.getParameter("selectedIndex");

      // Alternatively, get the specific RadioButton control
      // const oSelectedButton = oEvent.getSource().getSelectedButton();
      // const sText = oSelectedButton.getText();

      const OPSQuote1 = this.getView()?.byId("idOPSQuote1");
      const OPSQuote2 = this.getView()?.byId("idOPSQuote2");
      const OPSQuote3 = this.getView()?.byId("idOPSQuoteTaxInc");
      const OPSTaxInvc = this.getView()?.byId("idTaxSec");
      const QuotationSec = this.getView()?.byId("idQuotationSec");
      const idQuotation = this.getView()?.byId("idBtnQuotation");
      const idInvoice = this.getView()?.byId("idBtnInvoice");
      const chkBankDetail = this.getView()?.byId("chkBankDetail");
      const chkGST = this.getView()?.byId("chkGST");
      const chkTotal = this.getView()?.byId("chkTotal");
      // Logic based on selection
      if (iSelectedIndex === 0) {
        // Quotation
        OPSQuote1.setVisible(true);
        OPSQuote2.setVisible(true);
        QuotationSec.setVisible(true);
        OPSQuote3.setVisible(false);
        OPSTaxInvc.setVisible(false);
        idQuotation.setVisible(true);
        idInvoice.setVisible(false);
        chkBankDetail.setVisible(true);
        chkGST.setVisible(true);
        chkTotal.setVisible(true);
      } else {
        // TAX-INVOICE
        OPSQuote1.setVisible(false);
        OPSQuote2.setVisible(false);
        QuotationSec.setVisible(false);
        OPSQuote3.setVisible(true);
        OPSTaxInvc.setVisible(true);
        idQuotation.setVisible(false);
        idInvoice.setVisible(true);
        chkBankDetail.setVisible(false);
        chkGST.setVisible(false);
        chkTotal.setVisible(false);
      }
    },
    onTaxAddRow: function _onTaxAddRow() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/taxProducts");
      aProducts.push({
        taxpProductName: "",
        taxHSNCode: "",
        taxQuantity: 0,
        taxPrice: 0,
        taxTotal: "0.00"
      });
      oModel.setProperty("/taxProducts", aProducts);
    },
    onTaxDelete: function _onTaxDelete(oEvent) {
      // 1. Get the item (row) that was clicked
      const oItemToDelete = oEvent.getParameter("listItem");

      // 2. Get the binding context path (e.g., "/products/2")
      const sPath = oItemToDelete.getBindingContext().getPath();

      // 3. Extract the index from the path
      const iIndex = parseInt(sPath.split("/").pop());

      // 4. Get the model and the data array
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/taxProducts");

      // 5. Remove the element at the specific index
      aProducts.splice(iIndex, 1);

      // 6. Update the model to refresh the UI
      oModel.setProperty("/taxProducts", aProducts);
      this.onTaxCalc();
    },
    onTaxCalc: function _onTaxCalc() {
      const oModel = this.getView()?.getModel();
      const aProducts = oModel.getProperty("/taxProducts");
      let fGrandTotal = 0;
      let gst = 0;
      let taxcgst = 0;
      let taxsgst = 0;
      aProducts.forEach(oProduct => {
        // 1. Calculate individual row total
        const fQty = parseFloat(oProduct.taxQuantity) || 0;
        const fPrice = parseFloat(oProduct.taxPrice) || 0;
        let fRowTotal = 0;
        if (typeof fQty === "string") {
          fRowTotal = fPrice;
        } else if (typeof fQty === "number") {
          fRowTotal = fQty * fPrice;
        }

        // Update the row total in the model
        oProduct.taxTotal = fRowTotal.toFixed(2);

        // 2. Add to Grand Total
        taxcgst = taxcgst + fRowTotal * 0.09;
        taxsgst = taxsgst + fRowTotal * 0.09;
        fGrandTotal += fRowTotal;
      });

      // 3. Set the Grand Total property for the footer
      oModel.setProperty("/taxtotal", fGrandTotal.toFixed(2).toString());
      oModel.setProperty("/taxcgst", taxcgst.toFixed(2));
      oModel.setProperty("/taxsgst", taxsgst.toFixed(2));
      oModel.setProperty("/taxtotalSum", (fGrandTotal + taxcgst + taxsgst).toFixed(2).toString());

      // Refresh model to ensure UI updates
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

      // --- 0. SET PAGE BORDER ---
      // rect(x, y, width, height)
      doc.setDrawColor(0, 0, 0); // Black border
      doc.setLineWidth(0.3);
      doc.rect(margin, margin, pageWidth - margin * 2, pageHeight - margin * 2);

      // --- 1. Header Section ---
      doc.setFontSize(10);
      doc.text("TAX-INVOICE", 105, 10, {
        align: "center"
      });
      doc.text("Ph: 080-23481249", 170, 10);
      doc.line(startX, 12, endX, 12);
      doc.setFontSize(14);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(255, 0, 0);
      doc.text("IN-TELECOM SERVICES", 105, 20, {
        align: "center"
      });

      // doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(0, 0, 255);
      doc.text("#249, 7th Main, 4th Cross, 2nd Stage,Nagarabhavi, Bangalore-560072", 105, 25, {
        align: "center"
      });
      doc.line(startX, 32, endX, 32); // Horizontal Line Division

      // --- 2. Client & Invoice Info ---
      doc.setTextColor(0, 0, 0);
      doc.setFont("helvetica", "bold");
      doc.text("To,", 15, 42);
      doc.setFont("helvetica", "normal");

      // Billing Address (Left)
      const splitTo = doc.splitTextToSize(oData.taxHeader.To, 90);
      doc.text(splitTo, 15, 47);

      // Invoice Metadata (Right)
      const metaX = 130;
      doc.text(`GST No: ${oData.taxHeader.GSTNo}`, metaX, 42);
      doc.text(`Invoice No: ${oData.taxHeader.InvoiceNo}`, metaX, 48);
      doc.text(`Date: ${oData.taxHeader.Date}`, metaX, 54);
      doc.text(`PO No: ${oData.taxHeader.PONo}`, metaX, 60);
      doc.text(`PO Date: ${oData.taxHeader.PODate}`, metaX, 66);

      // --- 3. Line Items Table ---
      const tableRows = oData.taxProducts.map((item, index) => [index + 1, item.taxpProductName, item.taxHSNCode, Number(item.taxPrice).toFixed(2).toString(), item.taxQuantity, item.taxTotal]);

      //Totals and Calculations ---
      const totalAmount = oData.taxProducts.reduce((sum, item) => sum + parseFloat(item.taxTotal), 0);
      const taxVal = totalAmount * 0.09; // CGST/SGST 9%[span_13](end_span)
      const grandTotal = totalAmount + taxVal * 2;
      tableRows.push([{
        content: `TOTAL:`,
        colSpan: 5,
        styles: {
          halign: 'right',
          fontStyle: 'bold'
        }
      }, {
        content: totalAmount.toFixed(2),
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
        content: taxVal.toFixed(2),
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
        content: taxVal.toFixed(2),
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
        content: grandTotal.toFixed(2),
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
        // This function removes the vertical lines for the total rows
        didParseCell: function (data) {
          // Check if the current row is one of the total rows (the last 4 rows)
          const totalRowsCount = 4;
          const isTotalRow = data.row.index >= tableRows.length - totalRowsCount;
          if (isTotalRow) {
            // Remove vertical lines by setting lineWidth to 0 for sides
            // We only keep the top/bottom lines for the row structure
            data.cell.styles.lineWidth = {
              top: 0.1,
              right: 0,
              bottom: 0.1,
              left: 0
            };

            // For the very last cell (the amount), we can add the right border back to close the table
            if (data.column.index === 5) {
              data.cell.styles.lineWidth = {
                top: 0.1,
                right: 0.1,
                bottom: 0.1,
                left: 0
              };
            }
            // For the first cell, add the left border back
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
      const alignX = 190; // Fixed X-coordinate to align with the 'Amount' column end

      // doc.setFont("helvetica", "bold");

      // // Use { align: "right" } and a consistent X-coordinate to align them vertically
      // doc.text(`TOTAL: ${totalAmount.toFixed(2)}`, alignX, finalY, { align: "right" });
      // doc.line(startX, finalY + 2, endX, finalY + 2);
      // doc.text(`CGST @ 9%: ${taxVal.toFixed(2)}`, alignX, finalY + 6, { align: "right" });
      // doc.line(startX, finalY + 8, endX, finalY + 8);
      // doc.text(`SGST @ 9%: ${taxVal.toFixed(2)}`, alignX, finalY + 12, { align: "right" });
      // doc.line(startX, finalY + 14, endX, finalY + 14);
      // // Grand Total
      // doc.setFontSize(10);
      // doc.text(`GRAND TOTAL: ${grandTotal.toFixed(2)}`, alignX, finalY + 20, { align: "right" });

      // --- 5. Footer: Rupees, GST, and Bank ---
      doc.line(startX, finalY + 5, endX, finalY + 5); // Division Line

      doc.setFont("helvetica", "normal");
      const amountInWords = this.numberToWords(grandTotal);
      doc.text(`Rupees in words: ${amountInWords} Only`, 10, finalY + 10);
      doc.text(`Party GST No: ${oData.taxHeader.PartyGST || ""}`, 10, finalY + 18);
      doc.setFont("helvetica", "bold");
      doc.text("Bank Details", 10, finalY + 28);
      doc.setFont("helvetica", "normal");
      const splitBank = doc.splitTextToSize(oData.taxHeader.BankDetails, 100);
      doc.text(splitBank, 10, finalY + 33);

      // Signatory
      // doc.setFont("helvetica", "bold");
      // doc.text("For In-Telecom Services", 150, finalY + 63);
      // doc.text("Authorized Signatory", 150, finalY + 85);
      // SIGNATURE SECTION
      const sigY = doc.internal.pageSize.height - 40;
      if (this.sSignaturBase64) {
        doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, sigY, 70, 25);
      }
      window.open(doc.output("bloburl"), "_blank");
    },
    // Helper to convert number to Words
    numberToWords: function _numberToWords(num) {
      const a = ['', 'One ', 'Two ', 'Three ', 'Four ', 'Five ', 'Six ', 'Seven ', 'Eight ', 'Nine ', 'Ten ', 'Eleven ', 'Twelve ', 'Thirteen ', 'Fourteen ', 'Fifteen ', 'Sixteen ', 'Seventeen ', 'Eighteen ', 'Nineteen '];
      const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
      const regex = new RegExp('[0-9]{1,9}');
      const number = num.toString();
      if (!regex.test(number)) return '';
      if (num === 0) return 'Zero';

      // Split into segments for Crore, Lakh, Thousand, and Hundreds
      const n = ('000000000' + number).substr(-9).match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
      if (!n) return '';
      let str = '';
      // Use Number() to convert string indices to numbers for array access
      str += Number(n[1]) !== 0 ? (a[Number(n[1])] || b[Number(n[1][0])] + ' ' + a[Number(n[1][1])]) + 'Crore ' : '';
      str += Number(n[2]) !== 0 ? (a[Number(n[2])] || b[Number(n[2][0])] + ' ' + a[Number(n[2][1])]) + 'Lakh ' : '';
      str += Number(n[3]) !== 0 ? (a[Number(n[3])] || b[Number(n[3][0])] + ' ' + a[Number(n[3][1])]) + 'Thousand ' : '';
      str += Number(n[4]) !== 0 ? (a[Number(n[4])] || b[Number(n[4][0])] + ' ' + a[Number(n[4][1])]) + 'Hundred ' : '';
      str += Number(n[5]) !== 0 ? (str !== '' ? 'and ' : '') + (a[Number(n[5])] || b[Number(n[5][0])] + ' ' + a[Number(n[5][1])]) : '';
      return str.trim();
    }
  });
  return View;
});
//# sourceMappingURL=View-dbg.controller.js.map
