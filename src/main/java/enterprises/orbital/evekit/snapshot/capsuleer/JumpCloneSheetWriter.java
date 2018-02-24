package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.JumpClone;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class JumpCloneSheetWriter {

  // Singleton
  private JumpCloneSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // JumpClones.csv
    // JumpClonesMeta.csv
    stream.putNextEntry(new ZipEntry("JumpClones.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Jump Clone ID", "Location ID", "Clone Name", "Location Type");
    List<JumpClone> clones = CachedData.retrieveAll(at,
                                                    (contid, at1) -> JumpClone.accessQuery(acct, contid, 1000, false,
                                                                                           at1,
                                                                                           AttributeSelector.any(),
                                                                                           AttributeSelector.any(),
                                                                                           AttributeSelector.any(),
                                                                                           AttributeSelector.any()));

    for (JumpClone next : clones) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getJumpCloneID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getCloneName(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getLocationType(), SheetUtils.CellFormat.NO_STYLE));
      // @formatter:on
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("JumpClonesMeta.csv", stream, false, null);
    for (JumpClone next : clones) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "JumpClone");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
