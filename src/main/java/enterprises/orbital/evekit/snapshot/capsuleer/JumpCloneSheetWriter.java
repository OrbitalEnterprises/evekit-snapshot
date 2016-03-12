package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.JumpClone;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

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
    output.printRecord("ID", "Jump Clone ID", "Type ID", "Location ID", "Clone Name");
    List<JumpClone> clones = JumpClone.getAll(acct, at);

    for (JumpClone next : clones) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getJumpCloneID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getCloneName(), SheetUtils.CellFormat.NO_STYLE)); 
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
