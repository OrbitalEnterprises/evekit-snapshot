package enterprises.orbital.evekit.snapshot.common;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.common.Bookmark;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class BookmarkSheetWriter {

  // Singleton
  private BookmarkSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // Bookmarks.csv
    // BookmarksMeta.csv
    stream.putNextEntry(new ZipEntry("Bookmarks.csv"));
    final CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Folder ID", "Folder Name", "Folder Creator ID", "Bookmark ID", "Bookmark Creator ID",
                       "Created (Raw)", "Created", "Item ID",
                       "Type ID", "Location ID", "X", "Y", "Z", "Memo", "Note");

    List<Long> metaIDs = new ArrayList<>();
    CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
    CachedData.stream(at, (long contid, AttributeSelector ats) ->
                          Bookmark.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                               AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                               AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                               AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                               AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                               AttributeSelector.any()),
                      true, capture)
              .forEach(next -> {
                try {
                  // @formatter:off
                  SheetUtils.populateNextRow(output,
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getFolderID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getFolderName(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getFolderCreatorID(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getBookmarkID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getBookmarkCreatorID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getCreated(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getCreated()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getX(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(next.getY(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(next.getZ(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(next.getMemo(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getNote(), SheetUtils.CellFormat.NO_STYLE));
                  // @formatter:on
                } catch (IOException e) {
                  capture.handle(e);
                }
                metaIDs.add(next.getCid());
              });

    // Handle MetaData
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    CSVPrinter metaOutput = SheetUtils.prepForMetaData("BookmarksMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, metaOutput, next, "Bookmark");
      if (count > 0) output.println();
    }
    metaOutput.flush();
    stream.closeEntry();
  }

}
