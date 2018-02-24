package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.Implant;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

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
    List<Implant> implants = CachedData.retrieveAll(at,
                                                    (contid, at1) -> Implant.accessQuery(acct, contid, 1000, false, at1,
                                                                                         AttributeSelector.any()));

    for (Implant next : implants) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
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
