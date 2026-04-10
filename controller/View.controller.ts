import Controller from "sap/ui/core/mvc/Controller";
import JSONModel from "sap/ui/model/json/JSONModel";

declare var jspdf: any;

export default class View extends Controller {
    private sLogoBase64: string = "";



    public onInit(): void {
        const oData = {
            header: {
                To: "",
                Date: "",
                Location: "",
                Subject: "",
                Notes: "",
                TermsAndConditions: "",
                softwarePrereq: " "
            },
            products: [
                { productName: "", quantity: 0, price: 0, total: "" }
            ]
        };
        this.getView()?.setModel(new JSONModel(oData));
        this._loadLocalLogo("img/logo.jpg");
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

    public onAddRow(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/products");
        aProducts.push({ productName: "", price: 0, quantity: 1, total: "0.00" });
        oModel.setProperty("/products", aProducts);
    }

    public onCalc(): void {
        const oModel = this.getView()?.getModel() as JSONModel;
        const aProducts = oModel.getProperty("/products");
        aProducts.forEach((item: any) => {
            item.total = (parseFloat(item.price || 0) * parseFloat(item.quantity || 0)).toFixed(2);
        });
        oModel.refresh();
    }

    public onGeneratePDF(): void {
        debugger
        const jspdfLib = (window as any).jspdf;
        if (!jspdfLib) return;   

        const oModel = this.getView()?.getModel() as JSONModel;
        const oHeader = oModel.getProperty("/header");
        const aItems = oModel.getProperty("/products");
        const doc = new jspdfLib.jsPDF();
        const pageWidth = doc.internal.pageSize.width;

        // --- 1. HEADER SECTION ---
        // Increased Logo Size (Width: 50, Height: 25)
        if (this.sLogoBase64) {
            doc.addImage(this.sLogoBase64, 'JPEG', 14, 10, 70, 25);
        }

        doc.setTextColor(0, 0, 0);
        doc.setFontSize(9);
        doc.text("GST: 29AGKPP7288F1ZO", 14, 40); //[span_11](end_span)

        // Company Contact Info (Right Aligned)
        doc.setFontSize(10);
        doc.text("Ph: (Off.): 2348 1249", pageWidth - 14, 15, { align: 'right' });
        doc.text("97400 27266 / 98442 11193", pageWidth - 14, 20, { align: 'right' });
        doc.text("E-mail: intelecompatil@rediffmail.com", pageWidth - 14, 25, { align: 'right' });
        doc.setFontSize(9);
        doc.text("#249, 7th Main, 4th Cross, 2nd Stage,", pageWidth - 14, 30, { align: 'right' });
        doc.text("Nagarabhavi, Bangalore-560062", pageWidth - 14, 35, { align: 'right' });

        doc.line(14, 45, pageWidth - 14, 45);

        // --- 2. TO / SUB / DATE SECTION ---
        doc.setFont("helvetica", "bold");
        doc.text(`Date: ${oHeader.Date}`, pageWidth - 14, 55, { align: 'right' }); //[span_12](end_span)

        doc.text("To,", 14, 55);
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.To+","+oHeader.Location, 80), 14, 60); //[span_13](end_span)

        doc.setFont("helvetica", "bold");
        doc.text("Sub: " + oHeader.Subject, 14, 75); //[span_14](end_span)

        // --- 3. TABLE SECTION ---
        const tableBody = aItems.map((item: any, index: number) => [
            index + 1,
            item.productName,
            item.quantity,
            Number(item.price).toFixed(2).toString(),
            item.total
        ]);

        const subtotal = aItems.reduce((acc: number, cur: any) => acc + parseFloat(cur.total || 0), 0);

        (doc as any).autoTable({
            startY: 82,
             head: [['SI.No.', 'Particulars', 'Quantity', 'Rate', 'Total (Rs.)']], //[span_15](end_span)
            body: tableBody,
            theme: 'grid',
            headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], lineWidth: 0.1 },
            columnStyles: { 0: { cellWidth: 15 }, 1: { cellWidth: 90 }, 4: { halign: 'right' } }
        });

        let finalY = (doc as any).lastAutoTable.finalY;

        // Subtotal and GST Note
        doc.setFont("helvetica", "bold");
        doc.text("Total + (18% GST to be included)", 14, finalY + 10); //[span_16](end_span)
        doc.text(subtotal.toFixed(2), pageWidth - 14, finalY + 10, { align: 'right' });

        // --- 4. TERMS / NOTES / PREREQUISITES ---
        doc.setFontSize(9);
        doc.text("Terms & Conditions:", 14, finalY + 20); //[span_17](end_span)
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.TermsAndConditions, pageWidth - 28), 14, finalY + 26);
        // doc.text(doc.splitTextToSize(oHeader.TermsAndConditions, pageWidth - 28), 14, finalY + 42); //[span_18](end_span)

        doc.setFont("helvetica", "bold");
        doc.text("Notes:", 14, finalY + 60); //[span_19](end_span)
        doc.setFont("helvetica", "normal");
        doc.text(doc.splitTextToSize(oHeader.Notes, pageWidth - 28), 14, finalY + 65);

        // --- 5. SIGNATURE SECTION ---
        const sigY = doc.internal.pageSize.height - 40;
        doc.setFont("helvetica", "bold");
        doc.text("Yours faithfully", pageWidth - 70, sigY); //[span_20](end_span)
        doc.text("For IN-TELECOM SERVICE", pageWidth - 70, sigY + 5);
        doc.text("VR PATIL", pageWidth - 70, sigY + 20);
        doc.text("Senior Managing Executive", pageWidth - 70, sigY + 25);

        window.open(doc.output("bloburl"), "_blank");
    }
}