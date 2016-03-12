package enterprises.orbital.evekit.snapshot.common;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.Blueprint;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class BlueprintSheetWriter {

  // Singleton
  private BlueprintSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Blueprints.csv
    // BlueprintsMeta.csv
    stream.putNextEntry(new ZipEntry("Blueprints.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Item ID", "Location ID", "Type ID", "Type Name", "Flag ID", "Quantity", "Time Efficiency", "Material Efficiency", "Runs");
    List<Blueprint> blueprints = new ArrayList<Blueprint>();
    long contid = -1;
    List<Blueprint> batch = Blueprint.getAllBlueprints(acct, at, 1000, contid);
    while (batch.size() > 0) {
      blueprints.addAll(batch);
      contid = batch.get(batch.size() - 1).getItemID();
      batch = Blueprint.getAllBlueprints(acct, at, 1000, contid);
    }

    for (Blueprint next : blueprints) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTypeName(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getFlagID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getQuantity(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getTimeEfficiency(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getMaterialEfficiency(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getRuns(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
      // @formatter:on
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("BlueprintsMeta.csv", stream, false, null);
    for (Blueprint next : blueprints) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Blueprint");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
