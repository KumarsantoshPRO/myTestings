import Controller from "sap/ui/core/mvc/Controller";
import JSONModel from "sap/ui/model/json/JSONModel";
import MessageBox from "sap/m/MessageBox";
import CheckBox from "sap/m/CheckBox";

declare var jspdf: any;

export default class View extends Controller {
    private sLogoBase64: string = "";
    private sSignaturBase64: string = "";


    public onInit(): void {
        const oData = {
            header: {
                To: "",
                Date: "",
                Location: "",
                Subject: "",
                Notes: "",
                TermsAndConditions: "",
                BankDetails: "Payment Mode: Via Online\nBank: State Bank of India,\nBranch: Mallathahalli Branch\nName: In - Telecom Services\nC/A No: 64064045533\nIFSC Code: SBIN0040457"
            },
            products: [
                { productName: "", quantity: 0, price: 0, total: "" }
            ]
        };
        this.getView()?.setModel(new JSONModel(oData));
        this._loadLocalLogo("img/logo.jpg");
        this._loadSignature("img/Signature.jpg");
    }

    private _loadLocalLogo(sRelativePath: string): void {
        const sFullUrl = sap.ui.require.toUrl("my/app/generatebill/" + sRelativePath);
        const xhr = new XMLHttpRequest();
        xhr.onload = () => {
            const reader = new FileReader();
            reader.onloadend = () => {
                this.sLogoBase64 = reader.result as string;
            };
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
            reader.onloadend = () => {
                this.sSignaturBase64 = reader.result as string;
            };
            reader.readAsDataURL(xhr.response);
        };
        xhr.open("GET", sFullUrl);
        xhr.responseType = "blob";
        xhr.send();
    }

    public onAddRow(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/products");
        aProducts.push({ productName: "", price: 0, quantity: 1, total: "0.00" });
        oModel.setProperty("/products", aProducts);
    }

    public onDelete(oEvent: any): void {
        // 1. Get the item (row) that was clicked
        const oItemToDelete = oEvent.getParameter("listItem");

        // 2. Get the binding context path (e.g., "/products/2")
        const sPath = oItemToDelete.getBindingContext().getPath();

        // 3. Extract the index from the path
        const iIndex = parseInt(sPath.split("/").pop());

        // 4. Get the model and the data array
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/products");

        // 5. Remove the element at the specific index
        aProducts.splice(iIndex, 1);

        // 6. Update the model to refresh the UI
        oModel.setProperty("/products", aProducts);
        this.onCalc();
    }


    public onCalc(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/products");
        let fGrandTotal = 0;
        let gst = 0;
        let gstAmount = 0;

        aProducts.forEach((oProduct: any) => {
            // 1. Calculate individual row total
            const fQty = oProduct.quantity;
            const fPrice = parseFloat(oProduct.price) || 0;
            debugger;
            let fRowTotal = 0;
            if (typeof fQty === "string") {
                fRowTotal = fPrice;
            }
            else if (typeof fQty === "number") {
                fRowTotal = fQty * fPrice;
            }

            // Update the row total in the model
            oProduct.total = fRowTotal.toFixed(2);

            // 2. Add to Grand Total
            gstAmount = gstAmount + (fRowTotal * 0.18);
            fGrandTotal += fRowTotal;


        });

        // 3. Set the Grand Total property for the footer
        oModel.setProperty("/gst", gstAmount.toFixed(2));
        oModel.setProperty("/totalSum", (fGrandTotal + gstAmount).toFixed(2).toString());

        // Refresh model to ensure UI updates
        oModel.refresh();
    }
    public validateForm(): boolean {
        const oModel = this.getView()?.getModel() as JSONModel;
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
    }

    public onGeneratePDF(): void {
        if (this.validateForm()) {
            const jspdfLib = (window as any).jspdf;
            if (!jspdfLib) return;

            const oModel = this.getView()?.getModel() as JSONModel;
            const oHeader = oModel.getProperty("/header");
            const aItems = oModel.getProperty("/products");
            const doc = new jspdfLib.jsPDF();
            const pageWidth = doc.internal.pageSize.width;

            // HEADER SECTION
            if (this.sLogoBase64) {
                doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25);
            }

            doc.setTextColor(0, 0, 0);
            doc.setFontSize(9);

            // Company Contact Info (Right Aligned)
            doc.setFontSize(10);
            doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, { align: 'right' });
            doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, { align: 'right' });
            doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, { align: 'right' });
            doc.setFontSize(9);
            doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, { align: 'right' });
            doc.text("Nagarabhavi, Bangalore-560062", pageWidth - 14, 35, { align: 'right' });

            doc.line(14, 40, pageWidth - 14, 40);

            // TO / SUB / DATE SECTION 
            doc.setFont("helvetica", "bold");
            doc.text(`Date: ${oHeader.Date}`, pageWidth - 14, 55, { align: 'right' });

            doc.text("To,", 14, 55);
            doc.setFont("helvetica", "normal");
            doc.text(doc.splitTextToSize(oHeader.To + "\n" + oHeader.Location, 80), 14, 60);


            doc.setFont("helvetica", "bold");
            doc.text("Sub: " + oHeader.Subject, 14, 75);

            //TABLE SECTION
            const tableBody = aItems.map((item: any, index: number) => [
                index + 1,
                item.productName,
                item.quantity,
                Number(item.price).toFixed(2).toString(),
                item.total
            ]);

            const subtotal = aItems.reduce((acc: number, cur: any) => acc + parseFloat(cur.total || 0), 0);
            const gstAmount = Number(subtotal) * 0.18;
            const grandTotal = Number(subtotal) + Number(gstAmount);
            (doc as any).autoTable({
                startY: 82,
                head: [['SI.No.', 'Particulars', 'Quantity', 'Rate', 'Total (Rs.)']],
                body: tableBody,
                theme: 'grid',
                headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], lineWidth: 0.1 },
                columnStyles: { 0: { cellWidth: 15 }, 1: { cellWidth: 90 }, 4: { halign: 'right' } }
            });

            let finalY = (doc as any).lastAutoTable.finalY;

            // Subtotal and GST Note
            // Get the checkbox states
            const bShowGST = (this.getView()?.byId("chkGST") as CheckBox).getSelected();
            const bShowTotal = (this.getView()?.byId("chkTotal") as CheckBox).getSelected();

            doc.setFont("helvetica", "bold");
            // Conditional Logic for GST Amount
            if (bShowGST) {
                doc.text("18% GST Amount", 14, finalY + 10); //[span_16](end_span)
                doc.text(gstAmount.toFixed(2), pageWidth - 14, finalY + 10, { align: 'right' });
            }
            // Conditional Logic for Grand Total
            if (bShowTotal) {
                doc.text("Total", 14, finalY + 14); //[span_16](end_span)
                doc.text(grandTotal.toFixed(2), pageWidth - 14, finalY + 14, { align: 'right' });
            }

            // Terms & Conditions ---
            doc.setFontSize(9);
            doc.setFont("helvetica", "bold");
            doc.text("Terms & Conditions:", 14, finalY + 30);
            doc.setFont("helvetica", "normal");
            doc.text(doc.splitTextToSize(oHeader.TermsAndConditions, pageWidth - 28), 14, finalY + 36);

            // Bank Details ---
            const bBankDetails = (this.getView()?.byId("chkBankDetail") as CheckBox).getSelected();
            if (bBankDetails) {
                doc.setFontSize(9);
                doc.setFont("helvetica", "bold");
                doc.text("Bank Details:", pageWidth - 80, finalY + 30);
                doc.setFont("helvetica", "normal");
                doc.text(doc.splitTextToSize(oHeader.BankDetails, pageWidth - 80), pageWidth - 80, finalY + 36);
            }


            // Notes
            doc.setFont("helvetica", "bold");
            doc.text("Notes:", 14, finalY + 70);
            doc.setFont("helvetica", "normal");
            doc.text(doc.splitTextToSize(oHeader.Notes, pageWidth - 28), 14, finalY + 75);

            // SIGNATURE SECTION
            const sigY = doc.internal.pageSize.height - 45;

            if (this.sSignaturBase64) {
                doc.addImage(this.sSignaturBase64, 'JPEG', pageWidth - 80, sigY, 70, 25);
            }
            window.open(doc.output("bloburl"), "_blank");
        }
    }
}