package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.MailLabel;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class MailLabelSheetWriter {

  // Singleton
  private MailLabelSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // MailLabel.csv
    // MailLabelMeta.csv
    stream.putNextEntry(new ZipEntry("MailLabels.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Label ID", "Unread Count", "Name", "Color");
    List<Long> metaIDs = new ArrayList<>();
    List<MailLabel> lists = CachedData.retrieveAll(at,
                                                   (contid, at1) -> MailLabel.accessQuery(acct, contid, 1000, false,
                                                                                          at1,
                                                                                          AttributeSelector.any(),
                                                                                          AttributeSelector.any(),
                                                                                          AttributeSelector.any(),
                                                                                          AttributeSelector.any()));
    for (MailLabel next : lists) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getLabelID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getUnreadCount(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getName(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getColor(), SheetUtils.CellFormat.NO_STYLE));
      // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MailLabelsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "MailLabel");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
