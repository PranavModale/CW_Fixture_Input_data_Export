package Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItemRevision;

import Util.DatasetUtil;
import Util.ExcelUtil;

public class ExportFixtureToExcelService {

    private static final String TEMPLATE_DATASET_NAME = "Input Data Sheet for Fixture";
    private static final String OUTPUT_DATASET_NAME = "Input Data Sheet for Fixture.xlsx";
    private static final String ALLOWED_OBJECT_TYPE = "Fixture Assembly Revision";

    public static void execute() {

        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;

        try {
            InterfaceAIFComponent selected =
                    AIFUtility.getCurrentApplication().getTargetComponent();

            if (!(selected instanceof TCComponentItemRevision)) {
                showError(
                        "Export Fixture Data to Excel",
                        "Please select a Fixture Assembly Revision that contains the dataset \""
                                + TEMPLATE_DATASET_NAME + "\".");
                return;
            }

            TCComponentItemRevision itemRev = (TCComponentItemRevision) selected;

            String objectType = itemRev.getProperty("object_type");
            if (objectType == null || !ALLOWED_OBJECT_TYPE.equalsIgnoreCase(objectType.trim())) {
                showError(
                        "Export Fixture Data to Excel",
                        "Please select object type \"" + ALLOWED_OBJECT_TYPE + "\".");
                return;
            }

            TCComponentDataset dataset = DatasetUtil.findDataset(itemRev, TEMPLATE_DATASET_NAME);
            if (dataset == null) {
                showError(
                        "Export Fixture Data to Excel",
                        "Selected object does not contain the dataset \""
                                + TEMPLATE_DATASET_NAME
                                + "\".\n\nPlease select the correct revision.");
                return;
            }

            File templateFile = DatasetUtil.downloadDataset(dataset);

            fis = new FileInputStream(templateFile);
            workbook = new XSSFWorkbook(fis);
            fis.close();
            fis = null;

            Sheet sheet = workbook.getSheetAt(0);

            itemRev.refresh();

            ExcelUtil.processSheet(sheet, itemRev, getMapping());
            ExcelUtil.processTableBlocks(sheet, itemRev, getTableMapping());

            File outDir = new File("C:\\Temp");
            if (!outDir.exists()) {
                outDir.mkdirs();
            }

            File tempOutput = new File(
                    outDir,
                    "FixtureData_temp_" + System.currentTimeMillis() + ".xlsx");
            File finalOutput = new File(outDir, OUTPUT_DATASET_NAME);

            fos = new FileOutputStream(tempOutput);
            workbook.write(fos);
            fos.close();
            fos = null;

            workbook.close();
            workbook = null;

            Files.move(
                    tempOutput.toPath(),
                    finalOutput.toPath(),
                    StandardCopyOption.REPLACE_EXISTING);

            if (!templateFile.getAbsolutePath().equalsIgnoreCase(finalOutput.getAbsolutePath())) {
                try {
                    Files.deleteIfExists(templateFile.toPath());
                } catch (Exception ex) {
                    System.out.println(
                            "Warning: could not delete old downloaded template: "
                                    + templateFile.getAbsolutePath());
                }
            }

            DatasetUtil.uploadExcelToInformation(itemRev, finalOutput, OUTPUT_DATASET_NAME);

            try {
                Files.deleteIfExists(finalOutput.toPath());
            } catch (Exception ex) {
                System.out.println(
                        "Warning: uploaded successfully, but could not delete local file: "
                                + finalOutput.getAbsolutePath());
            }

            MessageDialog.openInformation(
                    getShell(),
                    "Export Fixture Data to Excel",
                    "Excel exported and uploaded successfully.");

        } catch (Exception e) {
            e.printStackTrace();
            showError(
                    "Export Fixture Data to Excel",
                    "Export failed.\n\n" + e.getMessage());

        } finally {
            try {
                if (fis != null) {
                    fis.close();
                }
            } catch (Exception ignored) {
            }

            try {
                if (fos != null) {
                    fos.close();
                }
            } catch (Exception ignored) {
            }

            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (Exception ignored) {
            }
        }
    }

    private static void showError(String title, String message) {
        MessageDialog.openError(getShell(), title, message);
    }

    private static Shell getShell() {
        if (Display.getDefault() == null) {
            return null;
        }
        return Display.getDefault().getActiveShell();
    }

    public static Map<String, String> getTableMapping() {
        return new HashMap<String, String>();
    }

    public static Map<String, String> getMapping() {

        Map<String, String> map = new HashMap<String, String>();

        // Input Data Sheet For Fixture
        map.put("Project No", "a2ProjectNo");
        map.put("Project Name", "a2CatalogueNoComp_P");
        map.put("Sub- Project No.", "a2SubProjectNo");
        map.put("Sub-Project Name", "a2SubProjectName");
        map.put("Manufacturer", "a2Manufacturer");
        map.put("Project Type", "a2ProjectType");
        map.put("Date", "a2Date");
        map.put("Prepared By", "a2PreparedBy");

        // PART
        map.put("Part Name", "a2PartNameTB");
        map.put("PLM ID/Drawing No. & Revision", "MULTI_PLM_DRAWING_REV");
        map.put("Version Change-Over(If any)", "a2VersionChangeOverNotePart");
        map.put("Additional Note (If any)", "a2AdditionalNotePart");
        map.put("Sub-Project Name", "a2ReferenceSubProjectName");
        map.put("Sub-Project No.", "a2ReferenceSubProjectNo");

        // Fixture
        map.put("Fixture Type", "a2FixtureType_New");
        map.put("Version Change Over Note", "a2VersionChangeOverNoteFix");
        map.put("Additional note (If any)", "a2AdditionalNoteFixture");
        map.put("Machine", "a2Machine");
        map.put("Std. Boughtout items for Fixture (eg. guiding elements,Pins, etc.)", "a2StdBoughtoutItems_Fixture");

        // Design (CAD Format)
        map.put("For In-house", "a2CADDesignFormatForInhouse");
        map.put("For Outsource", "a2CADDesignFormatOutsource");

        return map;
    }
}