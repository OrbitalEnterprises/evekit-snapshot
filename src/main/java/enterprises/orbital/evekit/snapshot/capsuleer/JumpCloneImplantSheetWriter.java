package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.JumpCloneImplant;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class JumpCloneImplantSheetWriter {

  // Singleton
  private JumpCloneImplantSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // JumpCloneImplants.csv
    // JumpCloneImplantsMeta.csv
    stream.putNextEntry(new ZipEntry("JumpCloneImplants.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Jump Clone ID", "Type ID");
    List<JumpCloneImplant> cloneImplants = CachedData.retrieveAll(at,
                                                                  (contid, at1) -> JumpCloneImplant.accessQuery(acct,
                                                                                                                contid,
                                                                                                                1000,
                                                                                                                false,
                                                                                                                at1,
                                                                                                                AttributeSelector.any(),
                                                                                                                AttributeSelector.any()));

    for (JumpCloneImplant next : cloneImplants) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getJumpCloneID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
      // @formatter:on
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("JumpCloneImplantsMeta.csv", stream, false, null);
    for (JumpCloneImplant next : cloneImplants) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "JumpCloneImplant");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
