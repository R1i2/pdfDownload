import { IInputs, IOutputs } from "./generated/ManifestTypes";
import {jsPDF} from "jspdf";
import * as autoTable  from "jspdf-autotable";
autoTable.applyPlugin(jsPDF);
export class PdfDownload implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    /**
     * Empty constructor.
     */
    declare private context:ComponentFramework.Context<IInputs>;
    declare private mainContainer:HTMLDivElement;
    declare private htmlText:string;
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        // Add control initialization code
        this.context = context;
        this.mainContainer = container;
        this.htmlText = `<table data-id="0" id="incident_auditTable"

    class="ms-Table dataTable cell-border hover stripe display compact is-striped table ui celled table-border"

    aria-describedby="incident_auditTable_info" style="width:100%;">

    <colgroup>

      <col data-dt-column="0" style="width: 0px;">

      <col data-dt-column="1" style="width: 0px;">

      <col data-dt-column="2" style="width: 0px;">

      <col data-dt-column="3" style="width: 0px;">

      <col data-dt-column="4" style="width: 0px;">

      <col data-dt-column="5" style="width: 0px;">

      <col data-dt-column="6" style="width: 0px;">

    </colgroup>

    <thead>

      <tr role="row">

        <th data-dt-column="0" rowspan="1" colspan="1" class="dt-orderable-asc dt-orderable-desc dt-ordering-asc"

          aria-sort="ascending" aria-label="Changed Date: Activate to invert sorting" tabindex="0"><span

            class="dt-column-title" role="button">Changed Date</span><span class="dt-column-order"></span></th>

        <th data-dt-column="1" rowspan="1" colspan="1" class="dt-orderable-asc dt-orderable-desc"

          aria-label="Operation: Activate to sort" tabindex="0"><span class="dt-column-title"

            role="button">Operation</span><span class="dt-column-order"></span></th>

        <th data-dt-column="2" rowspan="1" colspan="1" class="dt-orderable-asc dt-orderable-desc"

          aria-label="Modified By: Activate to sort" tabindex="0"><span class="dt-column-title" role="button">Modified

            By</span><span class="dt-column-order"></span></th>

        <th id="logicalheader" class="fixed-width-logical dt-orderable-asc dt-orderable-desc" data-dt-column="3"

          rowspan="1" colspan="1" aria-label="Logical Name: Activate to sort" tabindex="0"><span class="dt-column-title"

            role="button">Logical Name</span><span class="dt-column-order"></span></th>

        <th id="oldheader" class="fixed-width custom-td-width dt-orderable-asc dt-orderable-desc" data-dt-column="4"

          rowspan="1" colspan="1" aria-label="Old Value: Activate to sort" tabindex="0"><span class="dt-column-title"

            role="button">Old Value</span><span class="dt-column-order"></span></th>

        <th id="newheader" class="fixed-width custom-td-width dt-orderable-asc dt-orderable-desc" data-dt-column="5"

          rowspan="1" colspan="1" aria-label="New Value: Activate to sort" tabindex="0"><span class="dt-column-title"

            role="button">New Value</span><span class="dt-column-order"></span></th>

      </tr>

    </thead>

    <tbody>

      <tr id="3" data-id="fc7fe0ad-ec50-ee11-be6e-001dd80bf235" data-etn="incident">

        <td class="sorting_1">09/10/2024 7:25 AM</td>

        <td>Update</td>

        <td>Rishav Raj</td>

        <td class="logicalvalue fixed-width-logical">Case Status</td>

        <td class="oldvalue fixed-width">Open</td>

        <td class="newValue fixed-width">New</td>

        <td><span class="ms-Icon ms-Icon--Delete"></span></td>

      </tr>

      <tr id="0" data-id="fc7fe0ad-ec50-ee11-be6e-001dd80bf235" data-etn="incident">

        <td class="sorting_1">09/11/2023 2:46 PM</td>

        <td>Create</td>

        <td># ea-GDSP-SIS-DEV-SA</td>

        <td class="logicalvalue fixed-width-logical" style="white-space:nowrap">Dating Method Exam Date Flag<br>Height

          Unit<br>Number of Fetuses<br>AFP Billable?<br>Chorionicity ?<br>Fetus B:NT Unable to measure<br>Fetus A:NT

          Unable to measure<br>Fetal Demise &gt; 8 Weeks Flag<br>Fetus B:NT Measurement (mm)

          Flag<br>IsPatientBilGenerated<br>cfDNA Billable?<br>Fetus B CRL Measurement (mm)<br>CRL Measurement Flag<br>GA

          (Weeks)<br>GA Today<br>Was CRL Measured Flag<br>Non NT CRL Measurement (mm) Flag<br>Fetus B:NT CRL Unable to

          measure<br>GA (Days)<br>Fetus A:NT CRL Unable to measure<br>Owner<br>Fetus A:NT Measurement (mm) Flag<br>Do you

          want the Fetal Sex tested and reported<br>Fetal Sex from result<br>IVF Verification Flag<br>GA Dating Method CCC

          Verification Flag<br>Disclose Fetal Sex CCC Verified Flag<br>Estimated Due Date Flag<br>Latest Priority

          Specimen<br>Fetus A CRL Measurement (mm)<br>Uses Insulin<br>Fetus B:NT Unable to measure Flag<br>Fetus A:NT CRL

          Unable to measure Flag<br>Dating Method GA(Weeks)<br>Uses Heprin<br>Fetus A:NT CRL Measurement (mm)

          Flag<br>Fetus B:NT CRL Unable to measure Flag<br>GA (Weeks) Flag<br>Customer Contacted<br>Fetus B:NT CRL

          Measurement (mm) Flag<br>Fetal Demise &gt; 8 Weeks?<br>Ovum Donor Age at Donation Flag<br>Fetus B CRL

          Measurement (mm) Flag<br>Weight Unit<br>Fetus Count Verification Flag<br>GA (Days) Flag<br>CRL Unable To Measure

          FTA<br>CRL Unable To Measure FTB<br>Date at 24 Weeks Flag<br>Fetal Reduction<br>Ovum Donor Used for this

          Pregnancy?<br>Dating Method<br>Ovum Donor Age at Donation<br>Twins by NT Ultrasound Flag<br>Fetal Reduction

          Flag<br>Was CRL Measured?<br>Copilot Engaged<br>Diabetic Verification Flag<br>Case Status<br>Diabetic<br>Dating

          Method Exam Date<br>CCC<br>Fetus A:NT Unable to measure Flag<br>Fetal Sex<br>Date at 24 weeks</td>

        <td class="oldvalue fixed-width">

          null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null<br>null

        </td>

        <td class="newValue fixed-width">

          No<br>cm<br>1<br>Null<br>No<br>No<br>No<br>No<br>No<br>No<br>Null<br>null<br>No<br>20<br>No<br>No<br>No<br>No<br>0<br>No<br>CCC<br>No<br>No<br>N/A<br>No<br>No<br>No<br>No<br>P-23-251-71-757-33<br>null<br>No<br>No<br>No<br>No<br>No<br>No<br>No<br>No<br>No<br>No<br>No<br>No<br>No<br>Kg<br>No<br>No<br>No<br>No<br>No<br>No<br>null<br>Ultrasound<br>null<br>No<br>No<br>No<br>No<br>No<br>New<br>No<br>09/07/2023

          00:00:00<br>53-KAISER-GENETICS-AFP OFFICE-53 <br>No<br>Not Requested<br>10/05/2023 00:00:00</td>

        <td><span class="ms-Icon ms-Icon--Delete"></span></td>

      </tr>

      <tr id="1" data-id="fc7fe0ad-ec50-ee11-be6e-001dd80bf235" data-etn="incident">

        <td class="sorting_1">09/11/2023 2:48 PM</td>

        <td>Update</td>

        <td># ea-GDSP-SIS-Dev-WorkflowAdmin</td>

        <td class="logicalvalue fixed-width-logical">Age at term (Years)<br>Estimated Due Date</td>

        <td class="oldvalue fixed-width">null<br>null</td>

        <td class="newValue fixed-width">22.78<br>01/25/2024 00:00:00</td>

        <td><span class="ms-Icon ms-Icon--Delete"></span></td>

      </tr>

      <tr id="2" data-id="fc7fe0ad-ec50-ee11-be6e-001dd80bf235" data-etn="incident">

        <td class="sorting_1">09/11/2023 2:49 PM</td>

        <td>Update</td>

        <td>sisgdsp Admin1</td>

        <td class="logicalvalue fixed-width-logical">Case Status</td>

        <td class="oldvalue fixed-width">New</td>

        <td class="newValue fixed-width">Open</td>

        <td><span class="ms-Icon ms-Icon--Delete"></span></td>

      </tr>

    </tbody>

    <tfoot></tfoot>

  </table>`;
    }
    public generatePDF(){
        const doc:any = new jsPDF();
      
        // Get the HTML table element
        const table = document.getElementById('incident_auditTable');
      
        // Define column styles to control text wrapping
        const columnStyles = {
          0: { cellWidth: 30 }, // Changed Date
          1: { cellWidth: 20 }, // Operation
          2: { cellWidth: 30 }, // Modified By
          3: { cellWidth: 50 }, // Logical Name
          4: { cellWidth: 30 }, // Old Value
          5: { cellWidth: 30 }, // New Value
        };
      
        // Custom cell renderer to handle list items
        const customCellRenderer = (data:any) => {
           // Prevent default text rendering
        };
      
        // Add autotable to the document
        doc.autoTable({
          html: '#incident_auditTable',
          columnStyles: columnStyles,
          styles: {
            overflow: 'hidden', // Ensure text wraps within the cell
            cellPadding: 5, // Add padding to cells
          },
          theme: 'grid', // Use a grid theme for better alignment
          didParseCell: customCellRenderer, // Use custom cell renderer
        });
      
        // Save the PDF
        doc.save('table.pdf');
      }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Add code to update control view
        const downloadButton = document.createElement("button");
        downloadButton.id = "downloadButton";
        downloadButton.innerText = "Download"
        downloadButton.addEventListener("click",this.generatePDF);
        const tableContainer = document.createElement("div");
        tableContainer.innerHTML = this.htmlText;
        this.mainContainer.innerHTML = ``;
        this.mainContainer.appendChild(downloadButton);
        this.mainContainer.appendChild(tableContainer);
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }
}
