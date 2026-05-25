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
import TextArea from "sap/m/TextArea";
import Label from "sap/m/Label";
import VBox from "sap/m/VBox";
import HBox from "sap/m/HBox";
import Bar from "sap/m/Bar";
import Title from "sap/m/Title";
import SearchField from "sap/m/SearchField";
import Column from "sap/m/Column";
import Table from "sap/m/Table";
import Item from "sap/ui/core/Item";
import HTML from "sap/ui/core/HTML";
import Fragment from "sap/ui/core/Fragment";
import TableSelectDialog, { TableSelectDialog$ConfirmEvent, TableSelectDialog$SearchEvent } from "sap/m/TableSelectDialog";
import Filter from "sap/ui/model/Filter";
import FilterOperator from "sap/ui/model/FilterOperator";
import ListBinding from "sap/ui/model/ListBinding";
import ColumnListItem from "sap/m/ColumnListItem";
import DateFormat from "sap/ui/core/format/DateFormat";
import BusyIndicator from "sap/ui/core/BusyIndicator";

declare var jspdf: any;
declare var sap: any;
declare var XLSX: any;
declare var window: any;

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

    // Centralized Cloud Configuration Storage Pointers
    private sAnalyticsBinId: string = "6a1015ed6877513b27b3681c";
    private sBinId: string = "6a0ffb516610dd3ae8873764";
    private sCustomerBinId: string = "6a1415b26877513b27cc8043"; // Dedicated Customer Registry Cloud Database Bucket
    private sMasterKey: string = "$2a$10$QW2jbDsLe9nN3eAqSzg6v.Zz3jHv6WQfDk.HLgm4T3V0uvumxbx8i";

    // Storage Keys for Custom Template Tracking
    private sCustomNumKey: string = "my_billing_app_custom_numeric_seq";
    private sCustomSuffixKey: string = "my_billing_app_custom_suffix_format";
    private sCustomTriggerFlag: string = "my_billing_app_use_custom_start_flag";

    // Global Overlay Trackers for Automated Cascade Closures
    private oMainAnalyticsDialog: Dialog | null = null;


    public onInit(): void {
        let oData: any;
        const sSavedData = localStorage.getItem(this.sStorageKey);
        var oToday = new Date();
        var oDateFormat = (DateFormat as any).getInstance({ pattern: "dd-MM-yyyy" });
        var sTodayDate = oDateFormat.format(oToday);

        if (sSavedData) {
            try { oData = JSON.parse(sSavedData); } catch (e) { oData = null; }
        }

        if (!oData) {
            oData = {
                header: {
                    To: "", Date: sTodayDate, Subject: "", AddtionalInfo: "", Notes: "", TermsAndConditions: "", LeadSource: "Direct",
                    BankDetails: "Payment Mode: Via Online\nBank: State Bank of India,\nBranch: Mallathahalli Branch\nName: In - Telecom Services\nC/A No: 64064045533\nIFSC Code: SBIN0040457"
                },
                products: [{ productName: "", quantity: 1, price: "0.00", symbol: "", total: "0.00" }],
                taxHeader: {
                    To: "", GSTNo: "29AGKPP7288F1Z0", InvoiceNo: "", Date: sTodayDate, PONo: "", PODate: "", PartyGST: "", LeadSource: "Direct",
                    BankDetails: "Payment Mode: Via Online\nBank: State Bank of India,\nBranch: Mallathahalli Branch\nName: In - Telecom Services\nC/A No: 64064045533\nIFSC Code: SBIN0040457"
                },
                taxProducts: [{ taxpProductName: "", taxHSNCode: "", taxQuantity: 1, taxPrice: 0, taxSymbol: "", taxTotal: "0.00" }],
                cashHeader: {
                    cashTo: "", cashDate: sTodayDate, LeadSource: "Direct",
                    cashbankDetails: "Payment Mode: Via Online\nBank: State Bank of India,\nBranch: Mallathahalli Branch\nName: In - Telecom Services\nC/A No: 64064045533\nIFSC Code: SBIN0040457"
                },
                cashProducts: [{ cashBody: "", cashQuantity: "1", cashAmount: "0.00" }],
                gst: "0.00", totalSum: "0.00", taxtotal: "0.00", taxcgst: "0.00", taxsgst: "0.00", taxtotalSum: "0.00", cashTotalSum: "0.00"
            };
        }

        this.getView()?.setModel(new JSONModel(oData));
        this._loadLocalLogo("img/logo.jpg");
        this._loadSignature("img/Signature.jpg");
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
        if (!sAddress || sAddress.trim() === "") return "Document";
        return sAddress.split("\n")[0].replace(/[/\\?%*:|"<>\s]/g, "_").trim();
    }

    public async onGenerateNextInvoiceNumber(): Promise<void> {
        const oModel = this.getView()?.getModel() as JSONModel;
        BusyIndicator.show(0);

        try {
            let iNextSequence: number;
            let sSuffixFormat: string;
            const sIsCustomTriggerActive = localStorage.getItem(this.sCustomTriggerFlag);

            if (sIsCustomTriggerActive === "X") {
                iNextSequence = parseInt(localStorage.getItem(this.sCustomNumKey) || "1");
                sSuffixFormat = localStorage.getItem(this.sCustomSuffixKey) || `/${this._getDefaultFYLabel()}`;
                localStorage.removeItem(this.sCustomTriggerFlag);
            } else {
                const response = await fetch(`https://api.jsonbin.io/v3/b/${this.sBinId}/latest`, {
                    method: "GET",
                    headers: { "X-Master-Key": this.sMasterKey }
                });
                if (!response.ok) throw new Error("Could not retrieve current sequence from cloud server.");

                const result = await response.json();
                const iCurrentSequence = parseInt(result.record.counter) || 0;
                iNextSequence = iCurrentSequence + 1;
                sSuffixFormat = localStorage.getItem(this.sCustomSuffixKey) || `/${this._getDefaultFYLabel()}`;
            }

            const updateResponse = await fetch(`https://api.jsonbin.io/v3/b/${this.sBinId}`, {
                method: "PUT",
                headers: {
                    "Content-Type": "application/json",
                    "X-Master-Key": this.sMasterKey
                },
                body: JSON.stringify({ counter: iNextSequence })
            });
            if (!updateResponse.ok) throw new Error("Failed to write updated sequence index block back to database.");

            localStorage.setItem(this.sCustomNumKey, iNextSequence.toString());
            const sPaddedSequence = iNextSequence < 10 ? `0${iNextSequence}` : `${iNextSequence}`;
            const sFinalInvoiceNo = `${sPaddedSequence}${sSuffixFormat}`;

            oModel.setProperty("/taxHeader/InvoiceNo", sFinalInvoiceNo);
            MessageToast.show(`Global Invoice generated successfully: ${sFinalInvoiceNo}`);
        } catch (oError: any) {
            MessageBox.error(`Global Counter Error: ${oError.message || oError}`);
        } finally {
            BusyIndicator.hide();
        }
    }

    public onSetGlobalInvoiceSequence(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const oInput = new Input({ placeholder: "e.g., 99/26-27", width: "100%" });

        const oDialog = new Dialog({
            title: "Set Global Invoice Sequence Start Template",
            type: "Message",
            content: [new Text({ text: "Enter the custom starting sequence template matching your format (e.g., 99/26-27):" }), oInput],
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

                    BusyIndicator.show(0);
                    fetch(`https://api.jsonbin.io/v3/b/${this.sBinId}`, {
                        method: "PUT",
                        headers: {
                            "Content-Type": "application/json",
                            "X-Master-Key": this.sMasterKey
                        },
                        body: JSON.stringify({ counter: iNewSequenceValue })
                    }).then((res) => {
                        if (!res.ok) throw new Error("Database update rejected.");
                        localStorage.setItem(this.sCustomNumKey, iNewSequenceValue.toString());
                        localStorage.setItem(this.sCustomSuffixKey, sExtractedSuffix);
                        localStorage.setItem(this.sCustomTriggerFlag, "X");

                        oModel.setProperty("/taxHeader/InvoiceNo", "");
                        MessageToast.show(`Global sequence reset successfully to: ${sRawInput}`);
                        oDialog.close();
                    }).catch((err) => {
                        MessageBox.error(`Cloud Sync Failed: ${err.message}`);
                    }).finally(() => {
                        BusyIndicator.hide();
                    });
                }
            }),
            endButton: new Button({ text: "Cancel", press: () => { oDialog.close(); } }),
            afterClose: () => { oDialog.destroy(); }
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
        this.onSaveLocalStorage();

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

        this.onExcelDownload(null, "Quotation");
        this.onGenerateWord("Quotation");

        const sClient = this._getFirstLineName(oHeader.To);
        const sSelLead = oHeader.LeadSource || "Direct"; // Captured but excluded from final printed PDF outputs
        this._logDocumentAnalytics("Quotation", sTargetFilename, sClient, grandTotal, sSelLead);
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
        this.onSaveLocalStorage();
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

        this.onExcelDownload(null, "TAX-INVOICE");
        this.onGenerateWord("TAX-INVOICE");

        const sClient = this._getFirstLineName(oData.taxHeader.To);
        const sSelLead = oData.taxHeader.LeadSource || "Direct";
        this._logDocumentAnalytics("TAX-INVOICE", oData.taxHeader.InvoiceNo, sClient, grandTotal, sSelLead);
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
        aProducts.forEach((item: any) => { totalAmount += parseFloat(item.cashAmount || 0); });
        oModel.setProperty("/cashTotalSum", totalAmount.toFixed(2));
    }

    public onCashBillPDF(): void {
        this.onSaveLocalStorage();
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

        this.onExcelDownload(null, "Cash Bill");
        this.onGenerateWord("Cash Bill");

        const sClient = this._getFirstLineName(oData.cashHeader.cashTo);
        const sSelLead = oData.cashHeader.LeadSource || "Direct";
        this._logDocumentAnalytics("Cash Bill", sTargetFilename, sClient, totalAmount, sSelLead);
    }

    public onExcelDownload(oEvent: any, sOverrideMode?: string): void {
        const sSelectMode = sOverrideMode || (this.getView()?.byId("mySelect") as Select).getSelectedItem()?.getText();
        const oModel = this.getView()?.getModel() as JSONModel;
        let sFileName = "Export.xlsx";
        let headerData: any[] = [];
        let itemsData: any[] = [];

        headerData.push({ "FIELD": "MODE_METADATA", "VALUE": sSelectMode });

        if (sSelectMode === "Quotation") {
            sFileName = this._getFirstLineName(oModel.getProperty("/header/To")) + "_QuotationItems.xlsx";
            headerData.push({ "FIELD": "Header_To", "VALUE": oModel.getProperty("/header/To") });
            headerData.push({ "FIELD": "Header_Date", "VALUE": oModel.getProperty("/header/Date") });
            headerData.push({ "FIELD": "Header_Subject", "VALUE": oModel.getProperty("/header/Subject") });
            headerData.push({ "FIELD": "Header_AddtionalInfo", "VALUE": oModel.getProperty("/header/AddtionalInfo") });
            headerData.push({ "FIELD": "Header_Notes", "VALUE": oModel.getProperty("/header/Notes") });
            headerData.push({ "FIELD": "Header_TermsAndConditions", "VALUE": oModel.getProperty("/header/TermsAndConditions") });
            headerData.push({ "FIELD": "Header_BankDetails", "VALUE": oModel.getProperty("/header/BankDetails") });

            itemsData = oModel.getProperty("/products").map((item: any) => ({
                "Particulars": item.productName, "Rate": item.price, "Quantity": item.quantity, "Symbol": item.symbol, "Total": item.total
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

            itemsData = oModel.getProperty("/taxProducts").map((item: any) => ({
                "Particulars": item.taxpProductName, "HSN Code": item.taxHSNCode, "Rate": item.taxPrice, "Quantity": item.taxQuantity, "Symbol": item.taxSymbol, "Total": item.taxTotal
            }));
        } else {
            sFileName = this._getFirstLineName(oModel.getProperty("/cashHeader/cashTo")) + "_CashBillItems.xlsx";
            headerData.push({ "FIELD": "CashHeader_cashTo", "VALUE": oModel.getProperty("/cashHeader/cashTo") });
            headerData.push({ "FIELD": "CashHeader_cashDate", "VALUE": oModel.getProperty("/cashHeader/cashDate") });
            headerData.push({ "FIELD": "CashHeader_cashbankDetails", "VALUE": oModel.getProperty("/cashHeader/cashbankDetails") });

            itemsData = oModel.getProperty("/cashProducts").map((item: any) => ({
                "Particulars": item.cashBody, "Quantity": item.cashQuantity, "Amount": item.cashAmount
            }));
        }

        const ws = XLSX.utils.json_to_sheet(headerData);
        XLSX.utils.sheet_add_aoa(ws, [[]], { origin: -1 });
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
            const rawJsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[];
            const sSelectMode = (this.getView()?.byId("mySelect") as Select).getSelectedItem()?.getText();
            const oModel = this.getView()?.getModel() as JSONModel;

            const headerMap: any = {};
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
                const obj: any = {};
                tableHeaders.forEach((headerName: string, index: number) => {
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

                const parsedProducts = jsonDataItems.map((row: any) => ({
                    productName: row["Particulars"] || "", price: row["Rate"]?.toString() || "0.00",
                    quantity: parseInt(row["Quantity"]) || 1, symbol: row["Symbol"] || "", total: row["Total"]?.toString() || "0.00"
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

                const parsedTax = jsonDataItems.map((row: any) => ({
                    taxpProductName: row["Particulars"] || "", taxHSNCode: row["HSN Code"]?.toString() || "",
                    taxQuantity: parseInt(row["Quantity"]) || 1, taxPrice: parseFloat(row["Rate"]) || 0,
                    taxSymbol: row["Symbol"] || "", taxTotal: row["Total"]?.toString() || "0.00"
                }));
                oModel.setProperty("/taxProducts", parsedTax);
                this.onTaxCalc();
            } else {
                if (headerMap["CashHeader_cashTo"]) oModel.setProperty("/cashHeader/cashTo", headerMap["CashHeader_cashTo"]);
                if (headerMap["CashHeader_cashDate"]) oModel.setProperty("/cashHeader/cashDate", headerMap["CashHeader_cashDate"]);
                if (headerMap["CashHeader_cashbankDetails"]) oModel.setProperty("/cashHeader/cashbankDetails", headerMap["CashHeader_cashbankDetails"]);

                const parsedCash = jsonDataItems.map((row: any) => ({
                    cashBody: row["Particulars"] || "", cashQuantity: row["Quantity"]?.toString() || "1", cashAmount: row["Amount"]?.toString() || "0.00"
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

        if (Number(rupeePart) > 0) { result += convertWholeNumber(rupeePart); }
        else if (Number(paisaPart) > 0) { result += 'Zero'; }

        if (paisaPart && Number(paisaPart) > 0) {
            const p = paisaPart.match(/^(\d{2})$/);
            if (p) {
                const paisaWords = a[Number(p[1])] || b[Number(p[1][0])] + ' ' + a[Number(p[1][1])];
                result += (result !== '' ? ' and ' : '') + paisaWords.trim() + ' Paise';
            }
        }
        return result.trim();
    }

    public async onHSNValueHelpRequest(oEvent: any): Promise<void> {
        const oInput = oEvent.getSource() as Input;
        this.getView()?.data("valueHelpSourceInput", oInput);

        if (!this._pDialog) {
            this._pDialog = Fragment.load({
                id: this.getView()?.getId(),
                name: "my.app.generatebill.view.HSNValueHelpDialog",
                controller: this
            }) as Promise<TableSelectDialog>;

            const oDialog = await this._pDialog;
            this.getView()?.addDependent(oDialog);
        }

        const oDialog = await this._pDialog;
        const oBinding = oDialog.getBinding("items") as ListBinding;
        if (oBinding) oBinding.filter([]);
        oDialog.open("");
    }

    public onHSNValueHelpSearch(oEvent: TableSelectDialog$SearchEvent): void {
        const sValue = oEvent.getParameter("value");
        const oDialog = oEvent.getSource() as TableSelectDialog;
        const oBinding = oDialog.getBinding("items") as ListBinding;

        if (!sValue) {
            oBinding.filter([]);
            return;
        }

        const oFilterCode = new Filter("HSN_CD", FilterOperator.Contains, sValue);
        const oFilterDesc = new Filter("HSN_Description", FilterOperator.Contains, sValue);
        const oCombinedFilter = new Filter({ filters: [oFilterCode, oFilterDesc], and: false });
        oBinding.filter([oCombinedFilter]);
    }

    public onHSNValueHelpConfirm(oEvent: TableSelectDialog$ConfirmEvent): void {
        const oSelectedItem = oEvent.getParameter("selectedItem");
        const oInput = this.getView()?.data("valueHelpSourceInput") as Input;

        if (!oSelectedItem) return;

        const oContext = oSelectedItem.getBindingContext("hsnModel");
        if (oContext) {
            const sSelectedCode = oContext.getProperty("HSN_CD") as string;
            oInput.setValue(sSelectedCode);
        }
    }

    public async onShareToWhatsApp(): Promise<void> {
        const oView = this.getView();
        if (!oView) return;

        const sSelectMode = (oView.byId("mySelect") as Select).getSelectedItem()?.getText();
        const oModel = oView.getModel() as JSONModel;
        const jspdfLib = (window as any).jspdf;

        if (!jspdfLib) {
            MessageBox.error("jsPDF library is not loaded.");
            return;
        }

        let sClientName = ""; let sTotalAmount = ""; let sDocIdentifier = ""; let sRawAddressText = ""; let sTargetFilename = "Document.pdf";

        const doc = new jspdfLib.jsPDF();
        const pageWidth = doc.internal.pageSize.width; const pageHeight = doc.internal.pageSize.height; const margin = 5;

        doc.setDrawColor(0, 0, 0); doc.setLineWidth(0.3); doc.rect(margin, margin, pageWidth - (margin * 2), pageHeight - (margin * 2));

        if (this.sLogoBase64) { doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25); }

        doc.setFontSize(10);
        doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, { align: 'right' });
        doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, { align: 'right' });
        doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, { align: 'right' });
        doc.setFontSize(9);
        doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, { align: 'right' });
        doc.text("Nagarabhavi, Bangalore-560072", pageWidth - 14, 35, { align: 'right' });
        doc.line(5, 40, pageWidth - 5, 40);

        if (sSelectMode === "Quotation") {
            const oHeader = oModel.getProperty("/header");
            const aItems = oModel.getProperty("/products");
            sRawAddressText = oHeader.To || ""; sClientName = this._getFirstLineName(sRawAddressText);
            sTotalAmount = formatINR(oModel.getProperty("/totalSum")); sDocIdentifier = "Quotation"; sTargetFilename = `${sClientName}_Quotation.pdf`;

            doc.setFont("helvetica", "bold"); doc.text(`Date: ${oHeader.Date}`, pageWidth - 14, 45, { align: 'right' }); doc.text("To,", 14, 45);
            doc.setFont("helvetica", "normal"); doc.text(doc.splitTextToSize(oHeader.To, 80), 14, 50);
            doc.setFont("helvetica", "bold"); let finalHeaderY = 70; doc.text("Sub: " + oHeader.Subject, 14, 70);
            doc.setFont("helvetica", "normal"); if (oHeader.AddtionalInfo !== "") { doc.text(oHeader.AddtionalInfo, 14, finalHeaderY + 4); }

            const bShowGST = (oView.byId("chkGST") as CheckBox).getSelected();
            const bShowTotal = (oView.byId("chkTotal") as CheckBox).getSelected();
            const tableBody = aItems.map((item: any, index: number) => [index + 1, item.productName, item.quantity, item.price + item.symbol, formatINR(item.total)]);
            const subtotal = aItems.reduce((acc: number, cur: any) => acc + parseFloat(cur.total || 0), 0);
            const gstAmount = subtotal * 0.18; const grandTotal = subtotal + gstAmount;

            if (bShowGST) tableBody.push([{ content: '18% GST Amount', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold' } }, { content: formatINR(gstAmount), styles: { halign: 'right', fontStyle: 'bold' } }]);
            if (bShowTotal) tableBody.push([{ content: 'Total', colSpan: 4, styles: { halign: 'right', fontStyle: 'bold' } }, { content: formatINR(grandTotal), styles: { halign: 'right', fontStyle: 'bold' } }]);

            (doc as any).autoTable({
                startY: finalHeaderY + 8, head: [['Sl.No.', 'Particulars', 'Quantity', 'Rate', 'Total (Rs.)']], body: tableBody, theme: 'grid',
                styles: { fontSize: 9, lineColor: [0, 0, 0], lineWidth: 0.1 }, headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], fontStyle: 'bold' }
            });

            let finalY = (doc as any).lastAutoTable.finalY;
            if (oHeader.TermsAndConditions !== "") { finalY += 10; doc.text(doc.splitTextToSize(oHeader.TermsAndConditions, pageWidth - 28), 14, finalY + 4); }
            if ((oView.byId("chkBankDetail") as CheckBox).getSelected()) { finalY += 20; doc.text(doc.splitTextToSize(oHeader.BankDetails, pageWidth - 28), 14, finalY + 4); }

            const sSelLead = oHeader.LeadSource || "Direct";
            this._logDocumentAnalytics("Quotation", sTargetFilename, sClientName, grandTotal, sSelLead);
        } else if (sSelectMode === "TAX-INVOICE") {
            const oData = oModel.getData();
            sRawAddressText = oData.taxHeader.To || ""; sClientName = this._getFirstLineName(sRawAddressText);
            sTotalAmount = formatINR(oModel.getProperty("/taxtotalSum")); sDocIdentifier = `Tax Invoice (${oData.taxHeader.InvoiceNo || "N/A"})`; sTargetFilename = `${sClientName}_TaxInvoice.pdf`;

            doc.setFont("helvetica", "bold"); doc.text("To,", 15, 44); doc.setFont("helvetica", "normal"); doc.text(doc.splitTextToSize(oData.taxHeader.To, 90), 15, 49);
            const metaX = 130; doc.text(`GST No: ${oData.taxHeader.GSTNo}`, metaX, 44); doc.text(`Invoice No: ${oData.taxHeader.InvoiceNo}`, metaX, 50); doc.text(`Date: ${oData.taxHeader.Date}`, metaX, 56);

            const tableRows = oData.taxProducts.map((item: any, index: number) => [index + 1, item.taxpProductName, item.taxHSNCode, formatINR(item.taxPrice) + item.taxSymbol, item.taxQuantity, formatINR(item.taxTotal)]);
            const totalAmount = oData.taxProducts.reduce((sum: number, item: any) => sum + parseFloat(item.taxTotal), 0);
            const taxVal = totalAmount * 0.09; const grandTotal = totalAmount + (taxVal * 2);

            tableRows.push(
                [{ content: `TOTAL:`, colSpan: 5, styles: { halign: 'right', fontStyle: 'bold' } }, { content: formatINR(totalAmount), styles: { halign: 'right', fontStyle: 'bold' } }],
                [{ content: `GRAND TOTAL:`, colSpan: 5, styles: { halign: 'right', fontStyle: 'bold', fontSize: 10 } }, { content: formatINR(grandTotal), styles: { halign: 'right', fontStyle: 'bold', fontSize: 10 } }]
            );

            (doc as any).autoTable({ startY: 75, head: [["SI No.", "Particulars", "HSN Code", "Rate", "No.of Units", "Amount"]], body: tableRows, theme: 'grid' });
            const finalY = (doc as any).lastAutoTable.finalY + 5; doc.text(`Rupees in words: ${this.numberToWords(grandTotal)} Only`, 10, finalY + 10);

            const sSelLead = oData.taxHeader.LeadSource || "Direct";
            this._logDocumentAnalytics("TAX-INVOICE", oData.taxHeader.InvoiceNo, sClientName, grandTotal, sSelLead);
        } else {
            const oData = oModel.getData();
            sRawAddressText = oData.cashHeader.cashTo || ""; sClientName = this._getFirstLineName(sRawAddressText);
            sTotalAmount = formatINR(oModel.getProperty("/cashTotalSum")); sDocIdentifier = "Cash Bill"; sTargetFilename = `${sClientName}_CashBill.pdf`;

            doc.setFont("helvetica", "bold"); doc.text("To,", 15, 46); doc.setFont("helvetica", "normal"); doc.text(doc.splitTextToSize(oData.cashHeader.cashTo, 90), 15, 51);
            doc.text(`Date: ${oData.cashHeader.cashDate}`, pageWidth - 14, 46, { align: 'right' });

            const tableRows = oData.cashProducts.map((item: any, index: number) => [index + 1, item.cashBody, item.cashQuantity, formatINR(item.cashAmount)]);
            const totalAmount = oData.cashProducts.reduce((sum: number, item: any) => sum + parseFloat(item.cashAmount || 0), 0);
            tableRows.push([{ content: `GRAND TOTAL:`, colSpan: 3, styles: { halign: 'right', fontStyle: 'bold', fontSize: 10 } }, { content: formatINR(totalAmount), styles: { halign: 'right', fontStyle: 'bold', fontSize: 10 } }]);

            (doc as any).autoTable({ startY: 75, head: [["SI No.", "Particulars", "Quantity", "Amount"]], body: tableRows, theme: 'grid' });

            const sSelLead = oData.cashHeader.LeadSource || "Direct";
            this._logDocumentAnalytics("Cash Bill", sTargetFilename, sClientName, totalAmount, sSelLead);
        }

        if (this.sSignaturBase64) { doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, pageHeight - 40, 70, 25); }

        const blobHTML5 = doc.output('blob');
        const pdfFile = new File([blobHTML5], sTargetFilename, { type: 'application/pdf' });

        if (navigator.canShare && navigator.canShare({ files: [pdfFile] })) {
            try {
                await navigator.share({
                    files: [pdfFile], title: `${sDocIdentifier} - IN-TELECOM SERVICES`,
                    text: `Hello ${sClientName},\n\nYour *${sDocIdentifier}* from *IN-TELECOM SERVICES* is attached.\n\n*Amount Due:* Rs. ${sTotalAmount}\n\nThank you!`
                });
                MessageToast.show("Share prompt initialized successfully.");
            } catch (error) { MessageBox.error("Sharing processing was cancelled or encountered an error."); }
        } else {
            MessageBox.information(
                "Direct file attachment is not supported on this browser window configuration. The PDF file will be downloaded onto your local system shelf. You can drag and drop it into your WhatsApp target chat window manually.",
                {
                    title: "Browser Sharing Limitation",
                    onClose: () => {
                        doc.save(sTargetFilename);
                        const aPhoneMatch = sRawAddressText.match(/(?:(?:\+91)|91)?\s*([6-9]\d{9})\b/);
                        let sTargetPhone = aPhoneMatch ? "91" + aPhoneMatch[1] : "";
                        const sMessageText = `Hello ${sClientName},\n\nYour *${sDocIdentifier}* from *IN-TELECOM SERVICES* is ready.\n\n*Amount Due:* Rs. ${sTotalAmount}`;
                        sap.m.URLHelper.redirect(`https://api.whatsapp.com/send?phone=${sTargetPhone}&text=${encodeURIComponent(sMessageText)}`, true);
                    }
                }
            );
        }
    }

    public async onGenerateWord(sMode: string): Promise<void> {
        if (!(window as any).docxtemplater) {
            try {
                await new Promise<void>((resolve, reject) => {
                    const scriptZip = document.createElement("script");
                    scriptZip.src = "https://cdn.jsdelivr.net/npm/pizzip@3.1.4/dist/pizzip.min.js";
                    scriptZip.onload = () => {
                        const scriptDocx = document.createElement("script");
                        scriptDocx.src = "https://cdn.jsdelivr.net/npm/docxtemplater@3.45.0/build/docxtemplater.js";
                        scriptDocx.onload = () => resolve();
                        scriptDocx.onerror = (err) => reject(err);
                        document.head.appendChild(scriptDocx);
                    };
                    scriptZip.onerror = (err) => reject(err);
                    document.head.appendChild(scriptZip);
                });
            } catch (error) {
                MessageBox.error("Failed to dynamically load client-side template engines.");
                return;
            }
        }

        const PizZip = (window as any).PizZip;
        const Docxtemplater = (window as any).docxtemplater;
        const oModel = this.getView()?.getModel() as JSONModel;
        const oData = oModel.getData();

        let sTemplatePath = ""; let sTargetFilename = ""; let oTemplateDataMap: any = {};

        if (sMode === "Quotation") {
            sTemplatePath = "my/app/generatebill/model/Quotation_Template.docx";
            sTargetFilename = this._getFirstLineName(oData.header.To) + "_Quotation.docx";

            const subtotal = oData.products.reduce((acc: number, cur: any) => acc + parseFloat(cur.total || 0), 0);
            const gstAmount = subtotal * 0.18;
            oTemplateDataMap = {
                To: oData.header.To, Date: oData.header.Date, Subject: oData.header.Subject, AdditionalInfo: oData.header.AddtionalInfo,
                prds: oData.products.map((item: any, idx: number) => ({ i: idx + 1, productName: item.productName, quantity: item.quantity, price: item.price, symbol: item.symbol, total: formatINR(item.total) })),
                gstAmount: formatINR(gstAmount), grandTotal: formatINR(oModel.getProperty("/totalSum")), Notes: oData.header.Notes, TermsAndConditions: oData.header.TermsAndConditions
            };
        } else if (sMode === "TAX-INVOICE") {
            sTemplatePath = "my/app/generatebill/model/TaxInvoice_Template.docx";
            sTargetFilename = this._getFirstLineName(oData.taxHeader.To) + "_TaxInvoice.docx";

            const totalAmount = oData.taxProducts.reduce((sum: number, item: any) => sum + parseFloat(item.taxTotal), 0);
            const taxVal = totalAmount * 0.09; const grandTotal = totalAmount + (taxVal * 2);

            oTemplateDataMap = {
                To: oData.taxHeader.To, InvoiceNo: oData.taxHeader.InvoiceNo, Date: oData.taxHeader.Date, PONo: oData.taxHeader.PONo, PODate: oData.taxHeader.PODate,
                taxPrd: oData.taxProducts.map((item: any, idx: number) => ({ in: idx + 1, taxPart: item.taxpProductName, taxHSN: item.taxHSNCode, taxP: formatINR(item.taxPrice), taxS: item.taxSymbol, taxQ: item.taxQuantity, taxT: formatINR(item.taxTotal) })),
                total: formatINR(totalAmount), cgst: formatINR(taxVal), sgst: formatINR(taxVal), grandTotal: formatINR(grandTotal), words: this.numberToWords(grandTotal), PartyGST: oData.taxHeader.PartyGST
            };
        } else {
            sTemplatePath = "my/app/generatebill/model/CashBill_Template.docx";
            sTargetFilename = this._getFirstLineName(oData.cashHeader.cashTo) + "_CashBill.docx";

            const totalAmount = oData.cashProducts.reduce((sum: number, item: any) => sum + parseFloat(item.cashAmount || 0), 0);

            oTemplateDataMap = {
                cashTo: oData.cashHeader.cashTo, cashDate: oData.cashHeader.cashDate,
                cP: oData.cashProducts.map((item: any, idx: number) => ({ i: idx + 1, cashBody: item.cashBody, cashQuantity: item.cashQuantity, cashAmount: formatINR(item.cashAmount) })),
                cashTotalSum: formatINR(totalAmount), words: this.numberToWords(totalAmount)
            };
        }

        try {
            const sTemplateUrl = sap.ui.require.toUrl(sTemplatePath);
            const response = await fetch(sTemplateUrl);
            if (!response.ok) throw new Error("Could not fetch the specified Word template.");

            const arrayBuffer = await response.arrayBuffer();
            const zip = new PizZip(arrayBuffer);
            const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

            doc.setData(oTemplateDataMap);
            doc.render();

            const outBlob = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
            const link = document.createElement("a");
            link.href = URL.createObjectURL(outBlob); link.download = sTargetFilename; document.body.appendChild(link); link.click(); document.body.removeChild(link);

            MessageToast.show("Word document successfully downloaded from template mapping!");
        } catch (error: any) { MessageBox.error("Error filling out document template: " + error.message); }
    }

    /**
     * Asynchronously logs analytical metrics into the server container.
     */
    private async _logDocumentAnalytics(sDocType: string, sDocNo: string, sClient: string, fAmount: number, sLeadSource: string): Promise<void> {
        try {
            const oGetResponse = await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}/latest`, {
                method: "GET",
                headers: { "X-Master-Key": this.sMasterKey }
            });
            if (!oGetResponse.ok) throw new Error("Failed to read analytical historical logs.");
            const oGetResult = await oGetResponse.json();
            const aHistory = oGetResult.record.records || [];

            const oDate = new Date();
            const aMonths = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

            aHistory.push({
                id: sDocNo,
                type: sDocType,
                client: sClient,
                amount: fAmount,
                lead: sLeadSource, // Embedded parameter mapping
                timestamp: oDate.toISOString(),
                month: aMonths[oDate.getMonth()],
                year: oDate.getFullYear().toString()
            });

            await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}`, {
                method: "PUT",
                headers: {
                    "Content-Type": "application/json",
                    "X-Master-Key": this.sMasterKey
                },
                body: JSON.stringify({ records: aHistory })
            });
        } catch (oError) {
            console.error("Background analytical tracking logging engine failed: ", oError);
        }
    }

    // =========================================================================
    // REVENUE REFACTORING WORKSPACE INTERFACES (UPDATED CRUD LAYOUT DATA)
    // =========================================================================

    public async onShowAnalyticsGraph(): Promise<void> {
        const sTodayDateKey = new Date().toDateString();
        const sStoredDate = localStorage.getItem("analytics_view_date");
        let iDailyCount = parseInt(localStorage.getItem("analytics_daily_count") || "0", 10);

        if (sStoredDate === sTodayDateKey) {
            if (iDailyCount >= 50) {
                sap.m.MessageBox.error(`Daily View Limit Reached: You have opened the analytics dashboard 50 times today.`);
                return;
            }
            iDailyCount++;
            localStorage.setItem("analytics_daily_count", iDailyCount.toString());
        } else {
            localStorage.setItem("analytics_view_date", sTodayDateKey);
            localStorage.setItem("analytics_daily_count", "1");
        }

        sap.ui.core.BusyIndicator.show(0);

        try {
            const response = await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}/latest`, {
                method: "GET",
                headers: { "X-Master-Key": this.sMasterKey }
            });
            if (!response.ok) throw new Error("Could not retrieve analytical transaction entries.");

            const sLimitHeader = response.headers.get("x-ratelimit-limit");
            const sRemainingHeader = response.headers.get("x-ratelimit-remaining");

            if (sLimitHeader && sRemainingHeader) {
                const iTotalLimit = parseInt(sLimitHeader, 10);
                const iRemaining = parseInt(sRemainingHeader, 10);
                if ((iTotalLimit - iRemaining) >= 9000) {
                    sap.m.MessageBox.warning(`Critical API Consumption Warning: Monthly usage threshold exceeded.`);
                }
            }

            const result = await response.json();
            const aRecords: any[] = result.record.records || [];

            if (aRecords.length >= 60) {
                sap.m.MessageBox.information(`Storage Optimization Notice: Current ledger contains ${aRecords.length} entries.`);
            }

            const oHtmlMetricsDashboard = new sap.ui.core.HTML();
            const iCurrentYear = new Date().getFullYear();
            const sYearCurrent = iCurrentYear.toString();
            const sYearMinus1 = (iCurrentYear - 1).toString();
            const sYearMinus2 = (iCurrentYear - 2).toString();

            const oYearSelect = new sap.m.Select({

                selectedKey: sYearCurrent,
                items: [
                    new sap.ui.core.Item({ key: "ALL", text: "All Years" }),
                    new sap.ui.core.Item({ key: sYearMinus2, text: sYearMinus2 }),
                    new sap.ui.core.Item({ key: sYearMinus1, text: sYearMinus1 }),
                    new sap.ui.core.Item({ key: sYearCurrent, text: sYearCurrent })
                ]
            });

            const oMonthSelect = new sap.m.Select({

                selectedKey: "ALL",
                items: [
                    new sap.ui.core.Item({ key: "ALL", text: "All Months" }),
                    ...["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"].map(
                        sM => new sap.ui.core.Item({ key: sM, text: sM })
                    )
                ]
            });

            const oLeadSelect = new sap.m.Select({

                selectedKey: "ALL",
                items: [
                    new sap.ui.core.Item({ key: "ALL", text: "All Leads" }),
                    new sap.ui.core.Item({ key: "Justdial", text: "Justdial" }),
                    new sap.ui.core.Item({ key: "Sulekha", text: "Sulekha" }),
                    new sap.ui.core.Item({ key: "Direct", text: "Direct" })
                ]
            });

            const oFilterToolbar = new sap.m.HBox({
                width: "100%",
                justifyContent: "SpaceBetween",
                items: [oYearSelect, oMonthSelect, oLeadSelect]
            }).addStyleClass("sapUiSmallMarginBottom");

            const fnFilterAndRefreshDashboard = () => {
                const sSelYear = oYearSelect.getSelectedKey();
                const sSelMonth = oMonthSelect.getSelectedKey();
                const sSelLead = oLeadSelect.getSelectedKey();

                const aFiltered = aRecords.filter((rec: any) => {
                    const oDate = new Date(rec.timestamp);
                    const aMonths = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                    const sRowYear = rec.year || oDate.getFullYear().toString();
                    const sRowMonth = rec.month || aMonths[oDate.getMonth()];
                    const sRowLead = rec.lead || "Direct";

                    return (sSelYear === "ALL" || sRowYear === sSelYear) &&
                        (sSelMonth === "ALL" || sRowMonth === sSelMonth) &&
                        (sSelLead === "ALL" || sRowLead.toLowerCase() === sSelLead.toLowerCase());
                });

                let fQuoteTotal = 0, fTaxTotal = 0, fCashTotal = 0;
                aFiltered.forEach((rec: any) => {
                    if (rec.type === "Quotation") fQuoteTotal += rec.amount;
                    else if (rec.type === "TAX-INVOICE") fTaxTotal += rec.amount;
                    else if (rec.type === "Cash Bill") fCashTotal += rec.amount;
                });

                const fOverallRevenue = fQuoteTotal + fTaxTotal + fCashTotal;
                const getPercentageString = (fValue: number): string => fOverallRevenue === 0 ? "0%" : `${Math.min(100, Math.round((fValue / fOverallRevenue) * 100))}%`;

                oHtmlMetricsDashboard.setContent(`
                <div style="font-family: Arial, sans-serif; padding: 5px; min-width: 360px; color: #333;">
                    <div style="margin-bottom: 8px; font-size: 13px; font-weight: bold; color: #555;">
                        Filtered Operations count: <span style="color: #2b7d2b;">${aFiltered.length} Documents</span>
                    </div>
                    <div style="margin-bottom: 18px; font-size: 15px; font-weight: bold; color: #000; border-bottom: 2px solid #eee; padding-bottom: 6px;">
                        Gross Segment Volume: <span style="color: #0a6ed1;">Rs. ${formatINR(fOverallRevenue)}</span>
                    </div>
                    <div style="margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; font-size: 11px; margin-bottom: 4px; font-weight: bold;"><span>QUOTATIONS (Rs. ${formatINR(fQuoteTotal)})</span><span>${getPercentageString(fQuoteTotal)}</span></div>
                        <div style="width: 100%; background: #e0e0e0; border-radius: 4px; height: 14px; overflow: hidden;"><div style="width: ${getPercentageString(fQuoteTotal)}; background: #2b7d2b; height: 100%;"></div></div>
                    </div>
                    <div style="margin-bottom: 12px;">
                        <div style="display: flex; justify-content: space-between; font-size: 11px; margin-bottom: 4px; font-weight: bold;"><span>TAX INVOICES (Rs. ${formatINR(fTaxTotal)})</span><span>${getPercentageString(fTaxTotal)}</span></div>
                        <div style="width: 100%; background: #e0e0e0; border-radius: 4px; height: 14px; overflow: hidden;"><div style="width: ${getPercentageString(fTaxTotal)}; background: #e67e22; height: 100%;"></div></div>
                    </div>
                    <div style="margin-bottom: 4px;">
                        <div style="display: flex; justify-content: space-between; font-size: 11px; margin-bottom: 4px; font-weight: bold;"><span>CASH BILLS (Rs. ${formatINR(fCashTotal)})</span><span>${getPercentageString(fCashTotal)}</span></div>
                        <div style="width: 100%; background: #e0e0e0; border-radius: 4px; height: 14px; overflow: hidden;"><div style="width: ${getPercentageString(fCashTotal)}; background: #d32f2f; height: 100%;"></div></div>
                    </div>
                </div>
            `);
            };

            oYearSelect.attachChange(() => { fnFilterAndRefreshDashboard(); });
            oMonthSelect.attachChange(() => { fnFilterAndRefreshDashboard(); });
            oLeadSelect.attachChange(() => { fnFilterAndRefreshDashboard(); });
            fnFilterAndRefreshDashboard();

            this.oMainAnalyticsDialog = new Dialog({
                contentWidth: "440px",
                customHeader: new sap.m.Bar({
                    contentLeft: [new sap.m.Title({ text: "Business Revenue Analytics Dashboard" })],
                    contentRight: [new sap.m.Button({ icon: "sap-icon://decline", type: "Transparent", press: () => { this.oMainAnalyticsDialog?.close(); } })]
                }),
                content: [new sap.m.VBox({ items: [oFilterToolbar, oHtmlMetricsDashboard] }).addStyleClass("sapUiContentPadding")],
                buttons: [
                    new sap.m.Button({ text: "Download Summary", icon: "sap-icon://excel-attachment", type: "Accept", press: () => { this._downloadAnalyticsExcel(aRecords); } }),
                    new sap.m.Button({ text: "Edit History", type: "Reject", press: () => { this._verifyAdminPasswordBeforeManage(oYearSelect.getSelectedKey(), oMonthSelect.getSelectedKey()); } })
                ],
                afterClose: () => { this.oMainAnalyticsDialog?.destroy(); this.oMainAnalyticsDialog = null; }
            });
            this.oMainAnalyticsDialog.open();
        } catch (err: any) {
            sap.m.MessageBox.error(`Analytics Dashboard Fault: ${err.message || err}`);
        } finally {
            sap.ui.core.BusyIndicator.hide();
        }
    }

    private _verifyAdminPasswordBeforeManage(sTargetYear: string, sTargetMonth: string): void {
        const sMasterPassword = "clearMe@22";
        const oPasswordInput = new sap.m.Input({ type: sap.m.InputType.Password, placeholder: "Enter admin credentials", width: "100%" });

        const oSecurityDialog = new Dialog({
            title: "Security Verification",
            type: "Message",
            content: [new sap.m.Text({ text: "Please provide admin master authorization password:" }), oPasswordInput],
            beginButton: new sap.m.Button({
                text: "Verify",
                type: "Accept",
                press: () => {
                    if (oPasswordInput.getValue() !== sMasterPassword) {
                        sap.m.MessageBox.error("Authentication Failed!");
                        oPasswordInput.setValue("");
                        return;
                    }
                    oSecurityDialog.close();
                    this._openGranularDeletionWorkspace(sTargetYear, sTargetMonth);
                }
            }),
            endButton: new sap.m.Button({ text: "Abort", press: () => oSecurityDialog.close() }),
            afterClose: () => oSecurityDialog.destroy()
        });
        oSecurityDialog.open();
    }

    private async _openGranularDeletionWorkspace(sTargetYear: string, sTargetMonth: string): Promise<void> {
        sap.ui.core.BusyIndicator.show(0);
        try {
            const response = await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}/latest`, { method: "GET", headers: { "X-Master-Key": this.sMasterKey } });
            const result = await response.json();
            const aMasterRecords: any[] = result.record.records || [];

            const aFilteredRecords = aMasterRecords.filter((rec: any) => {
                const oDate = new Date(rec.timestamp);
                const aMonths = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                const sRowYear = rec.year || (isNaN(oDate.getTime()) ? "Legacy" : oDate.getFullYear().toString());
                const sRowMonth = rec.month || (isNaN(oDate.getTime()) ? "Legacy" : aMonths[oDate.getMonth()]);
                return (sTargetYear === "ALL" || sRowYear === sTargetYear) && (sTargetMonth === "ALL" || sRowMonth === sTargetMonth);
            });

            const aFormattedTableItems = aFilteredRecords.map((rec: any) => {
                let sFormattedIST = "N/A (Legacy)";
                if (rec.timestamp) {
                    const oDate = new Date(rec.timestamp);
                    if (!isNaN(oDate.getTime())) {
                        sFormattedIST = `${String(oDate.getDate()).padStart(2, '0')}-${String(oDate.getMonth() + 1).padStart(2, '0')}-${oDate.getFullYear()} ${String(oDate.getHours()).padStart(2, '0')}:${String(oDate.getMinutes()).padStart(2, '0')}:${String(oDate.getSeconds()).padStart(2, '0')}`;
                    }
                }
                return Object.assign({}, rec, { displayTimestamp: sFormattedIST, lead: rec.lead || "Legacy/Direct", isEditable: false });
            });

            const oSelectionModel = new sap.ui.model.json.JSONModel({ items: aFormattedTableItems });
            const oDeleteTable = new sap.m.Table({
                mode: sap.m.ListMode.MultiSelect,
                noDataText: "No operational tracking records found for this segment selection",
                columns: [
                    new Column({ header: new Label({ text: "Document ID", design: "Bold" }), width: "20%" }),
                    new Column({ header: new Label({ text: "Type", design: "Bold" }), width: "20%" }),
                    new Column({ header: new Label({ text: "Lead Source", design: "Bold" }), width: "20%" }),
                    new Column({ header: new Label({ text: "Amount (Rs.)", design: "Bold" }), width: "20%", hAlign: "Right" }),
                    new Column({ header: new Label({ text: "Date & Time", design: "Bold" }), width: "20%" })
                ]
            });

            oDeleteTable.bindItems({
                path: "/items",
                template: new ColumnListItem({
                    cells: [
                        new Text({ text: "{id}" }),
                        new VBox({ items: [new Text({ text: "{type}", visible: "{= !${isEditable} }" }), new Select({ selectedKey: "{type}", visible: "{isEditable}", width: "100%", items: [new Item({ key: "Quotation", text: "Quotation" }), new Item({ key: "TAX-INVOICE", text: "TAX-INVOICE" }), new Item({ key: "Cash Bill", text: "Cash Bill" })] })] }),
                        new VBox({ items: [new Text({ text: "{lead}", visible: "{= !${isEditable} }" }), new Select({ selectedKey: "{lead}", visible: "{isEditable}", width: "100%", items: [new Item({ key: "Justdial", text: "Justdial" }), new Item({ key: "Sulekha", text: "Sulekha" }), new Item({ key: "Direct", text: "Direct" }), new Item({ key: "Legacy/Direct", text: "Legacy/Direct" })] })] }),
                        new VBox({ items: [new Text({ text: { path: "amount", formatter: (val: any) => formatINR(val) }, visible: "{= !${isEditable} }" }), new Input({ value: "{amount}", type: sap.m.InputType.Number, visible: "{isEditable}" })] }),
                        new Text({ text: "{displayTimestamp}" })
                    ]
                })
            });
            oDeleteTable.setModel(oSelectionModel);

            const oWipeDialog = new Dialog({
                title: `Manage Segment Workspace: ${sTargetMonth}-${sTargetYear}`,
                contentWidth: "750px", contentHeight: "480px",
                content: [new VBox({ items: [new Text({ text: "Select checkbox elements below to perform manual adjustments or execute clearances:" }).addStyleClass("sapUiSmallMarginBottom"), oDeleteTable] }).addStyleClass("sapUiContentPadding")],
                buttons: [
                    new Button({
                        text: "Add New", icon: "sap-icon://add", type: "Accept",
                        press: () => { this._openNewLogFormInline(aMasterRecords, oSelectionModel, oWipeDialog); }
                    }),
                    new Button({
                        text: "Modify", icon: "sap-icon://edit", type: "Default",
                        press: () => {
                            const aSelected = oDeleteTable.getSelectedItems();
                            if (aSelected.length === 0) { sap.m.MessageToast.show("Please select one item."); return; }
                            aSelected.forEach((oItem: any) => oItem.getBindingContext().setProperty("isEditable", true));
                            oSelectionModel.refresh(true);
                        }
                    }),
                    new Button({
                        text: "Save", icon: "sap-icon://save", type: "Emphasized",
                        press: async () => {
                            sap.ui.core.BusyIndicator.show(0);
                            oSelectionModel.getProperty("/items").forEach((modifiedItem: any) => {
                                const oMatch = aMasterRecords.find((mRec: any) => mRec.id === modifiedItem.id);
                                if (oMatch) { oMatch.type = modifiedItem.type; oMatch.lead = modifiedItem.lead; oMatch.amount = parseFloat(modifiedItem.amount) || 0; }
                            });
                            try {
                                const response = await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}`, { method: "PUT", headers: { "Content-Type": "application/json", "X-Master-Key": this.sMasterKey }, body: JSON.stringify({ records: aMasterRecords }) });
                                if (!response.ok) throw new Error("Cloud update operation rejected.");
                                sap.m.MessageToast.show("Changes saved successfully to cloud!");
                                oWipeDialog.close();
                                if (this.oMainAnalyticsDialog) this.oMainAnalyticsDialog.close();
                            } catch (ex: any) { sap.m.MessageBox.error(ex.message); } finally { sap.ui.core.BusyIndicator.hide(); }
                        }
                    }),
                    new Button({
                        text: "Permanently Clear Selected", icon: "sap-icon://delete", type: "Reject",
                        press: () => {
                            const aSelectedUIItems = oDeleteTable.getSelectedItems();
                            if (aSelectedUIItems.length === 0) { sap.m.MessageToast.show("Please check at least one line item row box."); return; }
                            this._executeCloudArrayWipe(aSelectedUIItems, aMasterRecords, oWipeDialog);
                        }
                    }),
                    new Button({ text: "Cancel", press: () => oWipeDialog.close() })
                ],
                afterClose: () => oWipeDialog.destroy()
            });
            oWipeDialog.open();
        } catch (err: any) { sap.m.MessageBox.error(`Workspace Fault: ${err.message}`); } finally { sap.ui.core.BusyIndicator.hide(); }
    }

    private _openNewLogFormInline(aMasterRecords: any[], oWorkspaceModel: JSONModel, oWipeDialog: Dialog): void {
        const oIdIn = new Input({ placeholder: "e.g., INV-2026-089" });
        const oAmtIn = new Input({ type: sap.m.InputType.Number, placeholder: "Amount in Rs." });
        const oTypeSel = new Select({ items: [new Item({ key: "Quotation", text: "Quotation" }), new Item({ key: "TAX-INVOICE", text: "TAX-INVOICE" }), new Item({ key: "Cash Bill", text: "Cash Bill" })] });
        const oLeadSel = new Select({ items: [new Item({ key: "Justdial", text: "Justdial" }), new Item({ key: "Sulekha", text: "Sulekha" }), new Item({ key: "Direct", text: "Direct" }), new Item({ key: "Legacy/Direct", text: "Legacy/Direct" })] });

        const oAddLogDialog = new Dialog({
            title: "Append Manual Analytics Entry", contentWidth: "350px",
            content: [new VBox({ items: [new Label({ text: "Document ID*" }), oIdIn, new Label({ text: "Doc Type" }), oTypeSel, new Label({ text: "Lead Source" }), oLeadSel, new Label({ text: "Amount*" }), oAmtIn] }).addStyleClass("sapUiContentPadding")],
            buttons: [
                new Button({
                    text: "Insert Entry", type: "Accept", icon: "sap-icon://add",
                    press: async () => {
                        if (!oIdIn.getValue() || !oAmtIn.getValue()) { sap.m.MessageToast.show("Fill mandatory fields."); return; }
                        oAddLogDialog.close();
                        sap.ui.core.BusyIndicator.show(0);

                        const oNow = new Date();
                        const aMonths = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                        const oNewItem = {
                            id: oIdIn.getValue(), type: oTypeSel.getSelectedKey(), lead: oLeadSel.getSelectedKey(),
                            amount: parseFloat(oAmtIn.getValue()) || 0, client: "Manual Entry Profile",
                            timestamp: oNow.toISOString(), month: aMonths[oNow.getMonth()], year: oNow.getFullYear().toString()
                        };

                        aMasterRecords.push(oNewItem);
                        try {
                            await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}`, { method: "PUT", headers: { "Content-Type": "application/json", "X-Master-Key": this.sMasterKey }, body: JSON.stringify({ records: aMasterRecords }) });
                            sap.m.MessageToast.show("Manual ledger log saved.");
                            oWipeDialog.close();
                            if (this.oMainAnalyticsDialog) this.oMainAnalyticsDialog.close();
                        } catch (e: any) { sap.m.MessageBox.error(e.message); } finally { sap.ui.core.BusyIndicator.hide(); }
                    }
                }),
                new Button({ text: "Cancel", press: () => oAddLogDialog.close() })
            ],
            afterClose: () => oAddLogDialog.destroy()
        });
        oAddLogDialog.open();
    }

    private async _executeCloudArrayWipe(aSelectedUIItems: any[], aMasterRecords: any[], oWipeDialog: Dialog): Promise<void> {
        sap.ui.core.BusyIndicator.show(0);
        try {
            const aTargetIdsToDelete = aSelectedUIItems.map((oItem: any) => oItem.getBindingContext().getProperty("id"));
            const aCleanedRecords = aMasterRecords.filter((rec: any) => !aTargetIdsToDelete.includes(rec.id));

            const response = await fetch(`https://api.jsonbin.io/v3/b/${this.sAnalyticsBinId}`, {
                method: "PUT", headers: { "Content-Type": "application/json", "X-Master-Key": this.sMasterKey },
                body: JSON.stringify({ records: aCleanedRecords })
            });
            if (!response.ok) throw new Error("Cloud sync write parameters rejected updates.");

            sap.m.MessageToast.show(`Successfully erased ${aTargetIdsToDelete.length} record lines.`);
            oWipeDialog.close();
            if (this.oMainAnalyticsDialog) this.oMainAnalyticsDialog.close();
        } catch (err: any) {
            sap.m.MessageBox.error(`Wipe Fault: ${err.message || err}`);
        } finally {
            sap.ui.core.BusyIndicator.hide();
        }
    }

    private _downloadAnalyticsExcel(aRecords: any[]): void {
        if (!aRecords || aRecords.length === 0) { sap.m.MessageToast.show("No records available to export."); return; }
        const aExcelRows = aRecords.map((rec: any) => {
            let sFormattedIST = "";
            if (rec.timestamp) {
                const oDate = new Date(rec.timestamp);
                sFormattedIST = `${String(oDate.getDate()).padStart(2, '0')}-${String(oDate.getMonth() + 1).padStart(2, '0')}-${oDate.getFullYear()} ${String(oDate.getHours()).padStart(2, '0')}:${String(oDate.getMinutes()).padStart(2, '0')}:${String(oDate.getSeconds()).padStart(2, '0')}`;
            }
            return {
                "Document ID": rec.id, "Document Type": rec.type, "Lead Source": rec.lead || "Direct",
                "Client Name": rec.client, "Grand Total (Rs.)": rec.amount, "Date & Time (IST)": sFormattedIST
            };
        });

        const ws = XLSX.utils.json_to_sheet(aExcelRows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Revenue Overview");
        XLSX.writeFile(wb, "Business_Analytics_Summary.xlsx");
    }
    // =========================================================================
    // CUSTOMER PROFILE DIRECTORY (LOOKUP, CRUD & INJECTION FORM PANELS)
    // =========================================================================

    /**
     * Fetches custom records registry from jsonbin.io and initializes the 
     * single-select lookup table layout data grid workspace with fuzzy filtering.
     */
    /**
      * CUSTOMER DIRECTORY LOOKUP ENGINE
      * Dynamically instantiates the lookup catalog and handles runtime scoping states cleanly.
      */
    public async onOpenCustomerDirectory(): Promise<void> {
        sap.ui.core.BusyIndicator.show(0);
        try {
            const response = await fetch(`https://api.jsonbin.io/v3/b/${this.sCustomerBinId}/latest`, {
                method: "GET",
                headers: { "X-Master-Key": this.sMasterKey }
            });
            if (!response.ok) throw new Error("Could not retrieve customer registry database.");
            const result = await response.json();
            const aCustomers: any[] = result.record.customers || [];

            const oCustModel = new sap.ui.model.json.JSONModel({ customers: aCustomers });
            const oSearchField = new sap.m.SearchField({
                placeholder: "Filter customer listings cleanly...",
                width: "100%",
                liveChange: (oEvent: any) => {
                    const sQuery = oEvent.getParameter("newValue").toLowerCase();
                    const oBinding = oCustTable.getBinding("items");
                    const aFilters = [
                        new sap.ui.model.Filter("name", sap.ui.model.FilterOperator.Contains, sQuery),
                        new sap.ui.model.Filter("phone", sap.ui.model.FilterOperator.Contains, sQuery)
                    ];
                    oBinding.filter(sQuery ? new sap.ui.model.Filter({ filters: aFilters, and: false }) : []);
                }
            });

            const oCustTable = new sap.m.Table({
                mode: sap.m.ListMode.SingleSelectLeft,
                columns: [
                    new sap.m.Column({ header: new sap.m.Label({ text: "Profile Client Name", design: "Bold" }), width: "40%" }),
                    new sap.m.Column({ header: new sap.m.Label({ text: "Contact Details", design: "Bold" }), width: "30%" }),
                    new sap.m.Column({ header: new sap.m.Label({ text: "Customer GST No", design: "Bold" }), width: "30%" })
                ]
            });

            oCustTable.bindItems({
                path: "/customers",
                template: new sap.m.ColumnListItem({
                    cells: [
                        new sap.m.Text({ text: "{name}" }),
                        new sap.m.Text({ text: "{phone}" }),
                        new sap.m.Text({ text: "{gstin}" })
                    ]
                })
            });
            oCustTable.setModel(oCustModel);

            // FIX: We declare the layout instance shell first, allowing button bindings to resolve it safely later
            const oDirectoryDialog = new sap.m.Dialog({
                title: "Customer Registry",
                contentWidth: "620px",
                contentHeight: "500px",
                content: [oCustTable],
                afterClose: () => oDirectoryDialog.destroy()
            });

            // FIX: Build the header sub-elements dynamically AFTER the dialog pointer exists in memory
            const oCustomHeaderBar = new sap.m.Bar({
                contentLeft: [new sap.m.Title({ text: "Customer Registry" })],
                contentRight: [
                    new sap.m.Button({
                        icon: "sap-icon://decline",
                        type: "Transparent",
                        press: () => {
                            // Resolves safely now because the dialog instance is fully initialized in memory
                            oDirectoryDialog.close();
                        }
                    })
                ]
            });

            // Inject the customized components back into the dialog layout setup properties
            oDirectoryDialog.setCustomHeader(oCustomHeaderBar);
            oDirectoryDialog.setSubHeader(new sap.m.Bar({ contentLeft: [oSearchField] }));



            oDirectoryDialog.addButton(new sap.m.Button({
                text: "Edit", icon: "sap-icon://edit", type: "Attention",
                press: () => {
                    const oSelectedRow = oCustTable.getSelectedItem();
                    if (!oSelectedRow) { sap.m.MessageToast.show("Please select a record."); return; }
                    this._openCustomerFormWorkspace(aCustomers, oDirectoryDialog, oSelectedRow.getBindingContext().getObject());
                }
            }));

            oDirectoryDialog.addButton(new sap.m.Button({
                text: "Delete", icon: "sap-icon://delete", type: "Reject",
                press: () => {
                    const oSelectedRow = oCustTable.getSelectedItem();
                    if (!oSelectedRow) { sap.m.MessageToast.show("Please select a record."); return; }
                    this._deleteCustomerRecordDirect(aCustomers, oSelectedRow.getBindingContext().getObject(), oDirectoryDialog);
                }
            }));

            oDirectoryDialog.addButton(new sap.m.Button({
                text: "Register New", icon: "sap-icon://add", type: "Accept",
                press: () => { this._openCustomerFormWorkspace(aCustomers, oDirectoryDialog, null); }
            }));
            // Attach operational footprint buttons back onto the dialog base footer layout
            oDirectoryDialog.addButton(new sap.m.Button({
                text: "Inject", type: "Emphasized", icon: "sap-icon://accept",
                press: () => {
                    const oSelectedRow = oCustTable.getSelectedItem();
                    if (!oSelectedRow) { sap.m.MessageToast.show("Please select a record."); return; }
                    this._autoFillCustomerInputs(oSelectedRow.getBindingContext().getObject());
                    oDirectoryDialog.close();
                }
            }));

            oDirectoryDialog.open();
        } catch (err: any) {
            sap.m.MessageBox.error(`Directory Init Fault: ${err.message}`);
        } finally {
            sap.ui.core.BusyIndicator.hide();
        }
    }

    /**
     * Automatically maps database profile values into matching input form layout containers 
     * depending on the active visible view tab selection ("Quotation" | "TAX-INVOICE" | "Cash Bill").
     */
    private _autoFillCustomerInputs(oCust: any): void {
        const oView = this.getView();
        const sSelectMode = (oView?.byId("mySelect") as Select)?.getSelectedItem()?.getText() || "Quotation";
        const oModel = oView?.getModel() as JSONModel;
        if (sSelectMode === "Quotation") {
            oModel.setProperty("/header/To", oCust.name + "\n" + oCust.address);
        } else if (sSelectMode === "TAX-INVOICE") {
            oModel.setProperty("/taxHeader/To", oCust.name + "\n" + oCust.address);
            oModel.setProperty("/taxHeader/PartyGST", oCust.gstin);
        } else if (sSelectMode === "Cash Bill") {
            oModel.setProperty("/cashHeader/cashTo", oCust.name + "\n" + oCust.address);
        }
        sap.m.MessageToast.show(`Injected values matching profile: ${oCust.name}`);
        oModel.refresh(true);
    }

    /**
     * Sub-dialog form wrapper to handle both Registration (Add New) and Modification (Edit) tasks.
     */
    private _openCustomerFormWorkspace(aMasterList: any[], oParentLookupDialog: Dialog, oExistingCustToEdit: any | null): void {
        const bIsEditMode = !!oExistingCustToEdit;
        const oName = new sap.m.Input({ placeholder: "Company or Individual Profile Name", value: bIsEditMode ? oExistingCustToEdit.name : "" });
        const oPhone = new sap.m.Input({ placeholder: "Contact Phone number", value: bIsEditMode ? oExistingCustToEdit.phone : "" });
        const oGst = new sap.m.Input({ placeholder: "GST No", value: bIsEditMode ? oExistingCustToEdit.gstin : "" });
        const oAddr = new sap.m.TextArea({ placeholder: "Billing address maps", rows: 3, width: "100%", value: bIsEditMode ? oExistingCustToEdit.address : "" });

        const oFormDialog = new Dialog({
            title: bIsEditMode ? "Modify Client Profile" : "Register New Client Profile",
            contentWidth: "380px",
            content: [new sap.m.VBox({ items: [new sap.m.Label({ text: "Profile Customer Name*", design: "Bold" }), oName, new sap.m.Label({ text: "Contact Number Link*", design: "Bold" }), oPhone, new sap.m.Label({ text: "Customer GST No" }), oGst, new sap.m.Label({ text: "Street Level Address Details" }), oAddr] }).addStyleClass("sapUiContentPadding")],
            buttons: [
                new sap.m.Button({
                    text: bIsEditMode ? "Commit Edits" : "Save Profile", type: "Accept", icon: "sap-icon://save",
                    press: async () => {
                        if (!oName.getValue() || !oPhone.getValue()) { sap.m.MessageBox.error("Mandatory fields are missing variables."); return; }
                        oFormDialog.close();
                        sap.ui.core.BusyIndicator.show(0);

                        if (bIsEditMode) {
                            oExistingCustToEdit.name = oName.getValue();
                            oExistingCustToEdit.phone = oPhone.getValue();
                            oExistingCustToEdit.gstin = oGst.getValue();
                            oExistingCustToEdit.address = oAddr.getValue();
                        } else {
                            aMasterList.push({ name: oName.getValue(), phone: oPhone.getValue(), gstin: oGst.getValue(), address: oAddr.getValue() });
                        }

                        try {
                            await fetch(`https://api.jsonbin.io/v3/b/${this.sCustomerBinId}`, {
                                method: "PUT",
                                headers: { "Content-Type": "application/json", "X-Master-Key": this.sMasterKey },
                                body: JSON.stringify({ customers: aMasterList })
                            });
                            sap.m.MessageToast.show("Cloud profile synchronization sequence completed.");
                            oParentLookupDialog.close();
                        } catch (e: any) {
                            sap.m.MessageBox.error(e.message);
                        } finally {
                            sap.ui.core.BusyIndicator.hide();
                        }
                    }
                }),
                new sap.m.Button({
                    text: "Clear Fields", icon: "sap-icon://refresh",
                    press: () => { oName.setValue(""); oPhone.setValue(""); oGst.setValue(""); oAddr.setValue(""); }
                }),
                new sap.m.Button({ text: "Cancel Form", press: () => oFormDialog.close() })
            ],
            afterClose: () => oFormDialog.destroy()
        });
        oFormDialog.open();
    }

    /**
     * Excludes a selected customer profile entity from the master dataset array list 
     * and updates the repository cache backend synchronously.
     */
    private _deleteCustomerRecordDirect(aMasterList: any[], oCustTargetToDelete: any, oParentLookupDialog: Dialog): void {
        sap.m.MessageBox.confirm(`Wipe profile matching ${oCustTargetToDelete.name} permanently from the cloud repository directory registry?`, {
            title: "Confirm destructive operation",
            actions: [sap.m.MessageBox.Action.YES, sap.m.MessageBox.Action.NO],
            onClose: async (sAction: string) => {
                if (sAction !== sap.m.MessageBox.Action.YES) return;
                oParentLookupDialog.close();
                sap.ui.core.BusyIndicator.show(0);

                const aCleanedCustomers = aMasterList.filter((item: any) => item.phone !== oCustTargetToDelete.phone || item.name !== oCustTargetToDelete.name);
                try {
                    await fetch(`https://api.jsonbin.io/v3/b/${this.sCustomerBinId}`, {
                        method: "PUT",
                        headers: { "Content-Type": "application/json", "X-Master-Key": this.sMasterKey },
                        body: JSON.stringify({ customers: aCleanedCustomers })
                    });
                    sap.m.MessageToast.show("Erased profile record line.");
                } catch (e: any) {
                    sap.m.MessageBox.error(e.message);
                } finally {
                    sap.ui.core.BusyIndicator.hide();
                }
            }
        });
    }

}