package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.Implant;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class ImplantSheetWriter {

  // Singleton
  private ImplantSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Implants.csv
    // ImplantsMeta.csv
    stream.putNextEntry(new ZipEntry("Implants.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Type ID", "Type Name");
    List<Implant> implants = Implant.getAll(acct, at);

    for (Implant next : implants) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTypeName(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("ImplantsMeta.csv", stream, false, null);
    for (Implant next : implants) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Implant");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
