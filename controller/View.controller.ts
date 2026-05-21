import Controller from "sap/ui/core/mvc/Controller";
import JSONModel from "sap/ui/model/json/JSONModel";
import MessageBox from "sap/m/MessageBox";
import MessageToast from "sap/m/MessageToast";
import CheckBox from "sap/m/CheckBox";
import Select from "sap/m/Select";
import Button from "sap/m/Button";
import ObjectPageSection from "sap/uxap/ObjectPageSection";
import Dialog from "sap/m/Dialog";
import Input from "sap/m/Input";
import Text from "sap/m/Text";
import Fragment from "sap/ui/core/Fragment";
import TableSelectDialog, { TableSelectDialog$ConfirmEvent, TableSelectDialog$SearchEvent } from "sap/m/TableSelectDialog";
import Filter from "sap/ui/model/Filter";
import FilterOperator from "sap/ui/model/FilterOperator";
import ListBinding from "sap/ui/model/ListBinding";
import ColumnListItem from "sap/m/ColumnListItem";

declare var jspdf: any;
declare var XLSX: any; // External SheetJS Library Reference

const formatINR = (amount: number | string): string => {
    const value = typeof amount === 'string' ? parseFloat(amount) : amount;
    if (isNaN(value)) return '0.00';
    return new Intl.NumberFormat('en-IN', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
    }).format(value);
};

export default class View extends Controller {
    private _pDialog: Promise<TableSelectDialog> | null = null;
    private sLogoBase64: string = "";
    private sSignaturBase64: string = "";
    private sStorageKey: string = "my_billing_app_draft_data";
    private sSequenceKey: string = "my_billing_app_invoice_seq_counter";

    // Storage Keys for Custom Template Tracking
    private sCustomNumKey: string = "my_billing_app_custom_numeric_seq";
    private sCustomSuffixKey: string = "my_billing_app_custom_suffix_format";
    private sCustomTriggerFlag: string = "my_billing_app_use_custom_start_flag";

    public onInit(): void {
        let oData: any;
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
                products: [
                    { productName: "", quantity: 1, price: "0.00", symbol: "", total: "0.00" }
                ],
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
                taxProducts: [
                    { taxpProductName: "", taxHSNCode: "", taxQuantity: 1, taxPrice: 0, taxSymbol: "", taxTotal: "0.00" }
                ],
                cashHeader: {
                    cashTo: "",
                    cashDate: "",
                    cashbankDetails: "Payment Mode: Via Online\nBank: State Bank of India,\nBranch: Mallathahalli Branch\nName: In - Telecom Services\nC/A No: 64064045533\nIFSC Code: SBIN0040457"
                },
                cashProducts: [
                    { cashBody: "", cashQuantity: "1", cashAmount: "0.00" }
                ],
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
    }

    private _getDefaultFYLabel(): string {
        const oToday = new Date();
        const iCurrentMonth = oToday.getMonth();
        const iCurrentYear = oToday.getFullYear();
        let iStartFY = iCurrentMonth >= 3 ? iCurrentYear : iCurrentYear - 1;
        let iEndFY = iStartFY + 1;
        return `${iStartFY.toString().substring(2)}-${iEndFY.toString().substring(2)}`;
    }

    private _getFirstLineName(sAddress: string): string {
        if (!sAddress || sAddress.trim() === "") {
            return "Document";
        }
        return sAddress.split("\n")[0].replace(/[/\\?%*:|"<>\s]/g, "_").trim();
    }

    public onGenerateNextInvoiceNumber(): void {
        const oModel = this.getView()?.getModel() as JSONModel;

        let iNextSequence: number;
        let sSuffixFormat: string;

        const sIsCustomTriggerActive = localStorage.getItem(this.sCustomTriggerFlag);

        if (sIsCustomTriggerActive === "X") {
            iNextSequence = parseInt(localStorage.getItem(this.sCustomNumKey) || "1");
            sSuffixFormat = localStorage.getItem(this.sCustomSuffixKey) || `/${this._getDefaultFYLabel()}`;
            localStorage.removeItem(this.sCustomTriggerFlag);
        } else {
            let iLastSequence = parseInt(localStorage.getItem(this.sCustomNumKey) || "0");
            iNextSequence = iLastSequence + 1;
            sSuffixFormat = localStorage.getItem(this.sCustomSuffixKey) || `/${this._getDefaultFYLabel()}`;
        }

        localStorage.setItem(this.sCustomNumKey, iNextSequence.toString());

        const sPaddedSequence = iNextSequence < 10 ? `0${iNextSequence}` : `${iNextSequence}`;
        const sFinalInvoiceNo = `${sPaddedSequence}${sSuffixFormat}`;

        oModel.setProperty("/taxHeader/InvoiceNo", sFinalInvoiceNo);
        MessageToast.show(`Invoice generated successfully: ${sFinalInvoiceNo}`);
    }

    public onSetGlobalInvoiceSequence(): void {
        const oModel = this.getView()?.getModel() as JSONModel;

        // Create an input control dynamically
        const oInput = new Input({
            placeholder: "e.g., 99/26-27",
            width: "100%"
        });

        // Construct standard dialog container with an integrated text input field
        const oDialog = new Dialog({
            title: "Set Global Invoice Sequence Start Template",
            type: "Message",
            content: [
                new Text({ text: "Enter the custom starting sequence template matching your format (e.g., 99/26-27):" }),
                oInput
            ],
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

                    localStorage.setItem(this.sCustomNumKey, iNewSequenceValue.toString());
                    localStorage.setItem(this.sCustomSuffixKey, sExtractedSuffix);
                    localStorage.setItem(this.sCustomTriggerFlag, "X");

                    oModel.setProperty("/taxHeader/InvoiceNo", "");
                    MessageToast.show(`Next Invoice Number sequence configured to begin exactly at: ${sRawInput}`);

                    oDialog.close();
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
    }

    public onSaveLocalStorage(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        if (oModel) {
            localStorage.setItem(this.sStorageKey, JSON.stringify(oModel.getData()));
            MessageToast.show("Form data successfully saved locally!");
        } else {
            MessageBox.error("Error updating model context details.");
        }
    }

    public onClearLocalStorage(): void {
        MessageBox.confirm("Are you sure you want to clear your saved draft?", {
            actions: [MessageBox.Action.YES, MessageBox.Action.NO],
            onClose: (sAction: string | null) => {
                if (sAction === MessageBox.Action.YES) {
                    localStorage.removeItem(this.sStorageKey);
                    MessageToast.show("Saved data deleted. Refreshing page layout.");
                    this.onInit();
                }
            }
        });
    }

    private _loadLocalLogo(sRelativePath: string): void {
        const sFullUrl = sap.ui.require.toUrl("my/app/generatebill/" + sRelativePath);
        const xhr = new XMLHttpRequest();
        xhr.onload = () => {
            const reader = new FileReader();
            reader.onloadend = () => { this.sLogoBase64 = reader.result as string; };
            reader.readAsDataURL(xhr.response);
        };
        xhr.open("GET", sFullUrl);
        xhr.responseType = "blob";
        xhr.send();
    }

    private _loadSignature(sSigRelativePath: string): void {
        const sFullUrl = sap.ui.require.toUrl("my/app/generatebill/" + sSigRelativePath);
        const xhr = new XMLHttpRequest();
        xhr.onload = () => {
            const reader = new FileReader();
            reader.onloadend = () => { this.sSignaturBase64 = reader.result as string; };
            reader.readAsDataURL(xhr.response);
        };
        xhr.open("GET", sFullUrl);
        xhr.responseType = "blob";
        xhr.send();
    }

    public onSelectChange(oEvent: any): void {
        const iSelectedText = oEvent.getParameter("selectedItem").getText();
        this._updateSectionVisibilities(iSelectedText);
    }

    private _updateSectionVisibilities(sMode: string): void {
        const view = this.getView();
        if (!view) return;

        (view.byId("idOPSQuote1") as ObjectPageSection).setVisible(sMode === "Quotation");
        (view.byId("idOPSQuote2") as ObjectPageSection).setVisible(sMode === "Quotation");
        (view.byId("idQuotationSec") as ObjectPageSection).setVisible(sMode === "Quotation");
        (view.byId("idBtnQuotation") as Button).setVisible(sMode === "Quotation");
        (view.byId("chkBankDetail") as CheckBox).setVisible(sMode === "Quotation");
        (view.byId("chkGST") as CheckBox).setVisible(sMode === "Quotation");
        (view.byId("chkTotal") as CheckBox).setVisible(sMode === "Quotation");

        (view.byId("idOPSQuoteTaxInc") as ObjectPageSection).setVisible(sMode === "TAX-INVOICE");
        (view.byId("idTaxSec") as ObjectPageSection).setVisible(sMode === "TAX-INVOICE");
        (view.byId("idBtnInvoice") as Button).setVisible(sMode === "TAX-INVOICE");

        (view.byId("idOPSCash") as ObjectPageSection).setVisible(sMode === "Cash Bill");
        (view.byId("idCashSecTab") as ObjectPageSection).setVisible(sMode === "Cash Bill");
        (view.byId("idBtnCash") as Button).setVisible(sMode === "Cash Bill");
    }

    public onAddRow(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/products");
        aProducts.push({ productName: "", price: "0.00", symbol: "", quantity: 1, total: "0.00" });
        oModel.setProperty("/products", aProducts);
    }

    public onDelete(oEvent: any): void {
        const oItemToDelete = oEvent.getParameter("listItem") as ColumnListItem;
        const sPath = oItemToDelete.getBindingContext()!.getPath();
        const iIndex = parseInt(sPath.split("/").pop()!);
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/products");
        aProducts.splice(iIndex, 1);
        oModel.setProperty("/products", aProducts);
        this.onCalc();
    }

    public onCalc(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/products");
        let fGrandTotal = 0;
        let gstAmount = 0;

        aProducts.forEach((oProduct: any) => {
            const fQty = parseFloat(oProduct.quantity) || 0;
            const fPrice = parseFloat(oProduct.price) || 0;
            let fRowTotal = fQty * fPrice;
            oProduct.total = fRowTotal.toFixed(2);
            gstAmount += (fRowTotal * 0.18);
            fGrandTotal += fRowTotal;
        });

        oModel.setProperty("/gst", gstAmount.toFixed(2));
        oModel.setProperty("/totalSum", (fGrandTotal + gstAmount).toFixed(2));
        oModel.refresh();
    }

    public onGeneratePDF(): void {
        const jspdfLib = (window as any).jspdf;
        if (!jspdfLib) return;

        const oModel = this.getView()?.getModel() as JSONModel;
        const oHeader = oModel.getProperty("/header");
        const aItems = oModel.getProperty("/products");
        const doc = new jspdfLib.jsPDF();
        const pageWidth = doc.internal.pageSize.width;
        const pageHeight = doc.internal.pageSize.height;
        const margin = 5;

        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.3);
        doc.rect(margin, margin, pageWidth - (margin * 2), pageHeight - (margin * 2));

        if (this.sLogoBase64) { doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25); }

        doc.setFontSize(10);
        doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, { align: 'right' });
        doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, { align: 'right' });
        doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, { align: 'right' });
        doc.setFontSize(9);
        doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, { align: 'right' });
        doc.text("Nagarabhavi, Bangalore-560072", pageWidth - 14, 35, { align: 'right' });
        doc.line(5, 40, pageWidth - 5, 40);

        doc.setFont("helvetica", "bold");
        doc.text(`Date: ${oHeader.Date}`, pageWidth - 14, 45, { align: 'right' });
        doc.text("To,", 14, 45);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.To, 80), 14, 50);

        doc.setFont("helvetica", "bold");
        let finalHeaderY = 70;
        doc.text("Sub: " + oHeader.Subject, 14, 70);
        doc.setFont("helvetica", "normal");
        if (oHeader.AddtionalInfo !== "") { doc.text(oHeader.AddtionalInfo, 14, finalHeaderY + 4); }

        const bShowGST = (this.getView()?.byId("chkGST") as CheckBox).getSelected();
        const bShowTotal = (this.getView()?.byId("chkTotal") as CheckBox).getSelected();

        const tableBody = aItems.map((item: any, index: number) => [
            index + 1, item.productName, item.quantity, item.price + item.symbol, formatINR(item.total)
        ]);

        const subtotal = aItems.reduce((acc: number, cur: any) => acc + parseFloat(cur.total || 0), 0);
        const gstAmount = subtotal * 0.18;
        const grandTotal = subtotal + gstAmount;

        if (bShowGST) {
            tableBody.push([{ content: '18% GST Amount', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold' } }, { content: formatINR(gstAmount), styles: { halign: 'right', fontStyle: 'bold' } }]);
        }
        if (bShowTotal) {
            tableBody.push([{ content: 'Total', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold' } }, { content: formatINR(grandTotal), styles: { halign: 'right', fontStyle: 'bold' } }]);
        }

        (doc as any).autoTable({
            startY: finalHeaderY + 8,
            head: [['Sl.No.', 'Particulars', 'Quantity', 'Rate', 'Total (Rs.)']],
            body: tableBody,
            theme: 'grid',
            styles: { fontSize: 9, lineColor: [0, 0, 0], lineWidth: 0.1 },
            headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: 'bold' },
            columnStyles: { 0: { cellWidth: 15 }, 1: { cellWidth: 'auto' }, 2: { cellWidth: 20, halign: 'center' }, 3: { cellWidth: 25, halign: 'right' }, 4: { cellWidth: 30, halign: 'right' } }
        });

        let finalY = (doc as any).lastAutoTable.finalY;

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

        if ((this.getView()?.byId("chkBankDetail") as CheckBox).getSelected()) {
            finalY += 20;
            doc.setFont("helvetica", "bold");
            doc.text("Bank Details:", 14, finalY);
            doc.setFont("helvetica", "normal");
            doc.text(doc.splitTextToSize(oHeader.BankDetails, pageWidth - 28), 14, finalY + 4);
        }

        if (this.sSignaturBase64) { doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, pageHeight - 40, 70, 25); }

        const sTargetFilename = this._getFirstLineName(oHeader.To) + "_Quotation.pdf";
        doc.save(sTargetFilename);

        // Also trigger the matching Excel download
        this.onExcelDownload(null, "Quotation");
    }

    public onTaxAddRow(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/taxProducts");
        aProducts.push({ taxpProductName: "", taxHSNCode: "", taxQuantity: 1, taxPrice: 0, taxSymbol: "", taxTotal: "0.00" });
        oModel.setProperty("/taxProducts", aProducts);
    }

    public onTaxDelete(oEvent: any): void {
        const oItemToDelete = oEvent.getParameter("listItem") as ColumnListItem;
        const sPath = oItemToDelete.getBindingContext()!.getPath();
        const iIndex = parseInt(sPath.split("/").pop()!);
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/taxProducts");
        aProducts.splice(iIndex, 1);
        oModel.setProperty("/taxProducts", aProducts);
        this.onTaxCalc();
    }

    public onTaxCalc(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/taxProducts");
        let fGrandTotal = 0;
        let taxcgst = 0;
        let taxsgst = 0;

        aProducts.forEach((oProduct: any) => {
            const fQty = parseFloat(oProduct.taxQuantity) || 0;
            const fPrice = parseFloat(oProduct.taxPrice) || 0;
            let fRowTotal = fQty * fPrice;
            oProduct.taxTotal = fRowTotal.toFixed(2);
            taxcgst += (fRowTotal * 0.09);
            taxsgst += (fRowTotal * 0.09);
            fGrandTotal += fRowTotal;
        });

        oModel.setProperty("/taxtotal", fGrandTotal.toFixed(2));
        oModel.setProperty("/taxcgst", taxcgst.toFixed(2));
        oModel.setProperty("/taxsgst", taxsgst.toFixed(2));
        oModel.setProperty("/taxtotalSum", (fGrandTotal + taxcgst + taxsgst).toFixed(2));
        oModel.refresh();
    }

    public onTaxInvoicePDF(): void {
        const jspdfLib = (window as any).jspdf;
        if (!jspdfLib) return;
        const oModel = this.getView()?.getModel() as JSONModel;
        const oData = oModel.getData();
        const doc = new jspdfLib.jsPDF();
        const startX = 5;
        const endX = 205;
        const pageWidth = doc.internal.pageSize.width;
        const pageHeight = doc.internal.pageSize.height;
        const margin = 5;

        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.3);
        doc.rect(margin, margin, pageWidth - (margin * 2), pageHeight - (margin * 2));

        if (this.sLogoBase64) { doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25); }
        doc.setFontSize(10);
        doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, { align: 'right' });
        doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, { align: 'right' });
        doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, { align: 'right' });
        doc.setFontSize(9);
        doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, { align: 'right' });
        doc.text("Nagarabhavi, Bangalore-560072", pageWidth - 14, 35, { align: 'right' });
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

        const tableRows = oData.taxProducts.map((item: any, index: number) => [
            index + 1, item.taxpProductName, item.taxHSNCode, formatINR(item.taxPrice) + item.taxSymbol, item.taxQuantity, formatINR(item.taxTotal)
        ]);

        const totalAmount = oData.taxProducts.reduce((sum: number, item: any) => sum + parseFloat(item.taxTotal), 0);
        const taxVal = totalAmount * 0.09;
        const grandTotal = totalAmount + (taxVal * 2);

        tableRows.push(
            [{ content: `TOTAL:`, colSpan: 5, styles: { halign: 'right', fontStyle: 'bold' } }, { content: formatINR(totalAmount), styles: { halign: 'right', fontStyle: 'bold' } }],
            [{ content: `CGST @ 9%:`, colSpan: 5, styles: { halign: 'right' } }, { content: formatINR(taxVal), styles: { halign: 'right' } }],
            [{ content: `SGST @ 9%:`, colSpan: 5, styles: { halign: 'right' } }, { content: formatINR(taxVal), styles: { halign: 'right' } }],
            [{ content: `GRAND TOTAL:`, colSpan: 5, styles: { halign: 'right', fontStyle: 'bold', fontSize: 10 } }, { content: formatINR(grandTotal), styles: { halign: 'right', fontStyle: 'bold', fontSize: 10 } }]
        );

        (doc as any).autoTable({
            startY: 75,
            head: [["SI No.", "Particulars", "HSN Code", "Rate", "No.of Units", "Amount"]],
            body: tableRows,
            theme: 'grid',
            styles: { fontSize: 8, lineColor: [0, 0, 0], lineWidth: 0.1, textColor: [0, 0, 0] },
            headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center' },
            columnStyles: { 0: { halign: 'center' }, 3: { halign: 'right' }, 4: { halign: 'center' }, 5: { halign: 'right' } },
            didParseCell: function (data: any) {
                const totalRowsCount = 4;
                const isTotalRow = data.row.index >= (tableRows.length - totalRowsCount);
                if (isTotalRow) {
                    data.cell.styles.lineWidth = { top: 0.1, right: 0, bottom: 0.1, left: 0 };
                    if (data.column.index === 5) { data.cell.styles.lineWidth = { top: 0.1, right: 0.1, bottom: 0.1, left: 0 }; }
                    if (data.column.index === 0) { data.cell.styles.lineWidth = { top: 0.1, right: 0, bottom: 0.1, left: 0.1 }; }
                }
            }
        });

        const finalY = (doc as any).lastAutoTable.finalY + 5;
        doc.line(startX, finalY + 5, endX, finalY + 5);

        doc.setFont("helvetica", "normal");
        doc.text(`Rupees in words: ${this.numberToWords(grandTotal)} Only`, 10, finalY + 10);
        doc.text(`Party GST No: ${oData.taxHeader.PartyGST || ""}`, 10, finalY + 18);

        doc.setFont("helvetica", "bold");
        doc.text("Bank Details", 10, finalY + 28);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oData.taxHeader.BankDetails, 100), 10, finalY + 33);

        if (this.sSignaturBase64) { doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, pageHeight - 40, 70, 25); }

        const sTargetFilename = this._getFirstLineName(oData.taxHeader.To) + "_TaxInvoice.pdf";
        doc.save(sTargetFilename);

        // Also trigger the matching Excel download
        this.onExcelDownload(null, "TAX-INVOICE");
    }

    public onCashAddRow(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/cashProducts");
        aProducts.push({ cashBody: "", cashQuantity: "1", cashAmount: "0.00" });
        oModel.setProperty("/cashProducts", aProducts);
    }

    public onCashDelete(oEvent: any): void {
        const oItemToDelete = oEvent.getParameter("listItem") as ColumnListItem;
        const sPath = oItemToDelete.getBindingContext()!.getPath();
        const iIndex = parseInt(sPath.split("/").pop()!);
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/cashProducts");
        aProducts.splice(iIndex, 1);
        oModel.setProperty("/cashProducts", aProducts);
        this.onCashCalc();
    }

    public onCashCalc(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/cashProducts") || [];
        let totalAmount = 0;
        aProducts.forEach((item: any) => {
            totalAmount += parseFloat(item.cashAmount || 0);
        });
        oModel.setProperty("/cashTotalSum", totalAmount.toFixed(2));
    }

    public onCashBillPDF(): void {
        const jspdfLib = (window as any).jspdf;
        if (!jspdfLib) return;
        const oModel = this.getView()?.getModel() as JSONModel;
        const oData = oModel.getData();
        const doc = new jspdfLib.jsPDF();
        const startX = 5;
        const endX = 205;
        const pageWidth = doc.internal.pageSize.width;
        const pageHeight = doc.internal.pageSize.height;
        const margin = 5;

        doc.setDrawColor(0, 0, 0);
        doc.setLineWidth(0.3);
        doc.rect(margin, margin, pageWidth - (margin * 2), pageHeight - (margin * 2));

        if (this.sLogoBase64) { doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25); }
        doc.setFontSize(10);
        doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, { align: 'right' });
        doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, { align: 'right' });
        doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, { align: 'right' });
        doc.setFontSize(9);
        doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, { align: 'right' });
        doc.text("Nagarabhavi, Bangalore-560072", pageWidth - 14, 35, { align: 'right' });
        doc.line(5, 40, pageWidth - 5, 40);

        doc.setFont("helvetica", "bold");
        doc.text("To,", 15, 46);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oData.cashHeader.cashTo, 90), 15, 51);

        doc.text(`Date: ${oData.cashHeader.cashDate}`, pageWidth - 14, 46, { align: 'right' });
        doc.setFont("helvetica", "bold");
        doc.setFontSize(14);
        doc.text(`Cash Bill`, pageWidth / 2, 72, { align: 'center' });
        doc.setFont("helvetica", "normal");
        doc.setFontSize(10);

        const tableRows = oData.cashProducts.map((item: any, index: number) => [
            index + 1, item.cashBody, item.cashQuantity, formatINR(item.cashAmount)
        ]);

        const totalAmount = oData.cashProducts.reduce((sum: number, item: any) => sum + parseFloat(item.cashAmount || 0), 0);
        tableRows.push([
            { content: `GRAND TOTAL:`, colSpan: 3, styles: { halign: 'right', fontStyle: 'bold', fontSize: 10 } },
            { content: formatINR(totalAmount), styles: { halign: 'right', fontStyle: 'bold', fontSize: 10 } }
        ]);

        (doc as any).autoTable({
            startY: 75,
            head: [["SI No.", "Particulars", "Quantity", "Amount"]],
            body: tableRows,
            theme: 'grid',
            columnStyles: { 0: { cellWidth: 15 }, 1: { cellWidth: 'auto' }, 2: { cellWidth: 25, halign: 'center' }, 3: { cellWidth: 35, halign: 'right' } }
        });

        const finalY = (doc as any).lastAutoTable.finalY + 5;
        doc.line(startX, finalY + 5, endX, finalY + 5);

        doc.setFont("helvetica", "normal");
        doc.text(`Rupees in words: ${this.numberToWords(totalAmount)} Only`, 10, finalY + 10);

        doc.setFont("helvetica", "bold");
        doc.text("Bank Details", 10, finalY + 28);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oData.cashHeader.cashbankDetails, 100), 10, finalY + 33);

        if (this.sSignaturBase64) { doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, pageHeight - 40, 70, 25); }

        const sTargetFilename = this._getFirstLineName(oData.cashHeader.cashTo) + "_CashBill.pdf";
        doc.save(sTargetFilename);

        // Also trigger the matching Excel download
        this.onExcelDownload(null, "Cash Bill");
    }

    public onExcelDownload(oEvent: any, sOverrideMode?: string): void {
        const sSelectMode = sOverrideMode || (this.getView()?.byId("mySelect") as Select).getSelectedItem()?.getText();
        const oModel = this.getView()?.getModel() as JSONModel;
        let sFileName = "Export.xlsx";
        let headerData: any[] = [];
        let itemsData: any[] = [];

        // Add explicit Mode metadata row to help processing uploads safely
        headerData.push({ "FIELD": "MODE_METADATA", "VALUE": sSelectMode });

        if (sSelectMode === "Quotation") {
            sFileName = this._getFirstLineName(oModel.getProperty("/header/To")) + "_QuotationItems.xlsx";

            // Build key-value mapping rows for header parameters
            headerData.push({ "FIELD": "Header_To", "VALUE": oModel.getProperty("/header/To") });
            headerData.push({ "FIELD": "Header_Date", "VALUE": oModel.getProperty("/header/Date") });
            headerData.push({ "FIELD": "Header_Subject", "VALUE": oModel.getProperty("/header/Subject") });
            headerData.push({ "FIELD": "Header_AddtionalInfo", "VALUE": oModel.getProperty("/header/AddtionalInfo") });
            headerData.push({ "FIELD": "Header_Notes", "VALUE": oModel.getProperty("/header/Notes") });
            headerData.push({ "FIELD": "Header_TermsAndConditions", "VALUE": oModel.getProperty("/header/TermsAndConditions") });
            headerData.push({ "FIELD": "Header_BankDetails", "VALUE": oModel.getProperty("/header/BankDetails") });

            itemsData = oModel.getProperty("/products").map((item: any) => ({
                "Particulars": item.productName,
                "Rate": item.price,
                "Quantity": item.quantity,
                "Symbol": item.symbol,
                "Total": item.total
            }));
        } else if (sSelectMode === "TAX-INVOICE") {
            sFileName = this._getFirstLineName(oModel.getProperty("/taxHeader/To")) + "_TaxInvoiceItems.xlsx";

            headerData.push({ "FIELD": "TaxHeader_To", "VALUE": oModel.getProperty("/taxHeader/To") });
            headerData.push({ "FIELD": "TaxHeader_GSTNo", "VALUE": oModel.getProperty("/taxHeader/GSTNo") });
            headerData.push({ "FIELD": "TaxHeader_InvoiceNo", "VALUE": oModel.getProperty("/taxHeader/InvoiceNo") });
            headerData.push({ "FIELD": "TaxHeader_Date", "VALUE": oModel.getProperty("/taxHeader/Date") });
            headerData.push({ "FIELD": "TaxHeader_PONo", "VALUE": oModel.getProperty("/taxHeader/PONo") });
            headerData.push({ "FIELD": "TaxHeader_PODate", "VALUE": oModel.getProperty("/taxHeader/PODate") });
            headerData.push({ "FIELD": "TaxHeader_PartyGST", "VALUE": oModel.getProperty("/taxHeader/PartyGST") });
            // headerData.push({ "FIELD": "TaxHeader_BankDetails", "VALUE": oModel.getProperty("/taxHeader/BankDetails") });

            itemsData = oModel.getProperty("/taxProducts").map((item: any) => ({
                "Particulars": item.taxpProductName,
                "HSN Code": item.taxHSNCode,
                "Rate": item.taxPrice,
                "Quantity": item.taxQuantity,
                "Symbol": item.taxSymbol,
                "Total": item.taxTotal
            }));
        } else {
            sFileName = this._getFirstLineName(oModel.getProperty("/cashHeader/cashTo")) + "_CashBillItems.xlsx";

            headerData.push({ "FIELD": "CashHeader_cashTo", "VALUE": oModel.getProperty("/cashHeader/cashTo") });
            headerData.push({ "FIELD": "CashHeader_cashDate", "VALUE": oModel.getProperty("/cashHeader/cashDate") });
            headerData.push({ "FIELD": "CashHeader_cashbankDetails", "VALUE": oModel.getProperty("/cashHeader/cashbankDetails") });

            itemsData = oModel.getProperty("/cashProducts").map((item: any) => ({
                "Particulars": item.cashBody,
                "Quantity": item.cashQuantity,
                "Amount": item.cashAmount
            }));
        }

        // Generate worksheet starting with the custom structural metadata configurations
        const ws = XLSX.utils.json_to_sheet(headerData);

        // Append an empty row for separation space layout
        XLSX.utils.sheet_add_aoa(ws, [[]], { origin: -1 });

        // Append the operational Line Items directly beneath the spacing layer
        XLSX.utils.sheet_add_json(ws, itemsData, { origin: -1 });

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Items Data");
        XLSX.writeFile(wb, sFileName);
    }

    public onExcelUpload(oEvent: any): void {
        const oFileUploader = oEvent.getSource();
        const oFile = oEvent.getParameter("files")[0];
        if (!oFile) return;

        const reader = new FileReader();
        reader.onload = (e: any) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Raw JSON translation captures structural definitions natively
            const rawJsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[];
            const sSelectMode = (this.getView()?.byId("mySelect") as Select).getSelectedItem()?.getText();
            const oModel = this.getView()?.getModel() as JSONModel;

            // Differentiate header map items from structural matrix lines
            const headerMap: any = {};
            let lineItemsStartIndex = 0;

            for (let i = 0; i < rawJsonRows.length; i++) {
                const row = rawJsonRows[i];
                if (row && row[0] && typeof row[0] === 'string' && (row[0].includes("Header") || row[0] === "FIELD" || row[0] === "MODE_METADATA")) {
                    if (row[0] !== "FIELD") {
                        headerMap[row[0]] = row[1] || "";
                    }
                } else {
                    // Blank row spacing found or data rows began
                    lineItemsStartIndex = i;
                    break;
                }
            }

            // Slice out the remaining block data rows and convert using original object property mapping arrays
            const lineItemsRows = rawJsonRows.slice(lineItemsStartIndex).filter(row => row && row.length > 0);
            if (lineItemsRows.length === 0) return;

            // Extract table headers dynamically from array configuration parameters
            const tableHeaders = lineItemsRows[0];
            const jsonDataItems = lineItemsRows.slice(1).map(row => {
                const obj: any = {};
                tableHeaders.forEach((headerName: string, index: number) => {
                    obj[headerName] = row[index] !== undefined ? row[index] : "";
                });
                return obj;
            });

            if (sSelectMode === "Quotation") {
                // Restore Header Properties safely back into JSON Model path mappings
                if (headerMap["Header_To"]) oModel.setProperty("/header/To", headerMap["Header_To"]);
                if (headerMap["Header_Date"]) oModel.setProperty("/header/Date", headerMap["Header_Date"]);
                if (headerMap["Header_Subject"]) oModel.setProperty("/header/Subject", headerMap["Header_Subject"]);
                if (headerMap["Header_AddtionalInfo"]) oModel.setProperty("/header/AddtionalInfo", headerMap["Header_AddtionalInfo"]);
                if (headerMap["Header_Notes"]) oModel.setProperty("/header/Notes", headerMap["Header_Notes"]);
                if (headerMap["Header_TermsAndConditions"]) oModel.setProperty("/header/TermsAndConditions", headerMap["Header_TermsAndConditions"]);
                if (headerMap["Header_BankDetails"]) oModel.setProperty("/header/BankDetails", headerMap["Header_BankDetails"]);

                const parsedProducts = jsonDataItems.map((row: any) => ({
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
                // if (headerMap["TaxHeader_BankDetails"]) oModel.setProperty ("/taxHeader/BankDetails", headerMap["TaxHeader_BankDetails"]);

                const parsedTax = jsonDataItems.map((row: any) => ({
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

                const parsedCash = jsonDataItems.map((row: any) => ({
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
    }

   

    private numberToWords(num: number): string {
        const a = ['', 'One ', 'Two ', 'Three ', 'Four ', 'Five ', 'Six ', 'Seven ', 'Eight ', 'Nine ', 'Ten ', 'Eleven ', 'Twelve ', 'Thirteen ', 'Fourteen ', 'Fifteen ', 'Sixteen ', 'Seventeen ', 'Eighteen ', 'Nineteen '];
        const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

        if (num === 0) return 'Zero';

        const convertWholeNumber = (amountStr: string): string => {
            const n = ('000000000' + amountStr).substring(amountStr.length + 9 - 9).match(/^(\d{2})(\d{2})(\d{2})(\d{1})(\d{2})$/);
            if (!n) return '';

            let str = '';
            str += (Number(n[1]) !== 0) ? (a[Number(n[1])] || b[Number(n[1][0])] + ' ' + a[Number(n[1][1])]) + 'Crore ' : '';
            str += (Number(n[2]) !== 0) ? (a[Number(n[2])] || b[Number(n[2][0])] + ' ' + a[Number(n[2][1])]) + 'Lakh ' : '';
            str += (Number(n[3]) !== 0) ? (a[Number(n[3])] || b[Number(n[3][0])] + ' ' + a[Number(n[3][1])]) + 'Thousand ' : '';
            str += (Number(n[4]) !== 0) ? (a[Number(n[4])] || b[Number(n[4][0])] + ' ' + a[Number(n[4][1])]) + 'Hundred ' : '';
            str += (Number(n[5]) !== 0) ? ((str !== '') ? 'and ' : '') + (a[Number(n[5])] || b[Number(n[5][0])] + ' ' + a[Number(n[5][1])]) : '';
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

    /**
     * Triggered when the value help icon inside the Input field is clicked
     */
    public async onHSNValueHelpRequest(oEvent: any): Promise<void> {
        // Store reference to the source input control to set the value later
        const oInput = oEvent.getSource() as Input;
        this.getView()?.data("valueHelpSourceInput", oInput);

        // Load the XML Fragment asynchronously
        if (!this._pDialog) {
            this._pDialog = Fragment.load({
                id: this.getView()?.getId(),
                name: "my.app.generatebill.view.HSNValueHelpDialog",
                controller: this
            }) as Promise<TableSelectDialog>;

            // Connect dialog to the lifecycle of the view
            const oDialog = await this._pDialog;
            this.getView()?.addDependent(oDialog);
        }

        const oDialog = await this._pDialog;

        // Ensure all rows are shown initially by resetting filters on open
        const oBinding = oDialog.getBinding("items") as ListBinding;
        if (oBinding) {
            oBinding.filter([]);
        }

        oDialog.open("");
    }

    /**
     * Handles search option for all columns (HSN_CD and HSN_Description)
     */
    public onHSNValueHelpSearch(oEvent: TableSelectDialog$SearchEvent): void {
        const sValue = oEvent.getParameter("value");
        const oDialog = oEvent.getSource() as TableSelectDialog;
        const oBinding = oDialog.getBinding("items") as ListBinding;

        if (!sValue) {
            oBinding.filter([]);
            return;
        }

        // Create individual filters matching case-insensitive substrings
        const oFilterCode = new Filter("HSN_CD", FilterOperator.Contains, sValue);
        const oFilterDesc = new Filter("HSN_Description", FilterOperator.Contains, sValue);

        // Combine using 'false' flag for OR operator to query both columns simultaneously
        const oCombinedFilter = new Filter({
            filters: [oFilterCode, oFilterDesc],
            and: false
        });

        oBinding.filter([oCombinedFilter]);
    }

    /**
     * Triggered when a row is selected or clicked
     */
    public onHSNValueHelpConfirm(oEvent: TableSelectDialog$ConfirmEvent): void {
        const oSelectedItem = oEvent.getParameter("selectedItem");
        const oInput = this.getView()?.data("valueHelpSourceInput") as Input;

        if (!oSelectedItem) {
            return;
        }

        // Get the structural bound context object properties
        const oContext = oSelectedItem.getBindingContext("hsnModel");
        if (oContext) {
            const sSelectedCode = oContext.getProperty("HSN_CD") as string;

            // Populate value to your Input element
            oInput.setValue(sSelectedCode);
        }
    }
}
