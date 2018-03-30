package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.MemberRoleHistory;
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

public class MemberRoleHistorySheetWriter {

  // Singleton
  private MemberRoleHistorySheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // MemberRoleHistory.csv
    // MemberRoleHistoryMeta.csv
    stream.putNextEntry(new ZipEntry("MemberRoleHistory.csv"));
    final CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Character ID", "Changed At (Raw)", "Changed At", "Issuer ID", "Role Type", "Role Name",
                       "Old?");
    List<Long> metaIDs = new ArrayList<>();

    CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
    CachedData.stream(at, (long contid, AttributeSelector ats) ->
                          MemberRoleHistory.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                                        AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                        AttributeSelector.any(), AttributeSelector.any()),
                      true, capture)
              .forEach(next -> {
                try {
                  // @formatter:off
                  SheetUtils.populateNextRow(output,
                                             new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getChangedAt(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(new Date(next.getChangedAt()), SheetUtils.CellFormat.DATE_STYLE),
                                             new DumpCell(next.getIssuerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getRoleType(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getRoleName(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.isOld(), SheetUtils.CellFormat.NO_STYLE));
                  // @formatter:on
                } catch (IOException e) {
                  capture.handle(e);
                }
                metaIDs.add(next.getCid());
              });

    output.flush();
    stream.closeEntry();

    // Handle MetaData
    CSVPrinter metaOutput = SheetUtils.prepForMetaData("MemberRoleHistoryMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, metaOutput, next, "MemberRoleHistory");
      if (count > 0) metaOutput.println();
    }
    metaOutput.flush();
    stream.closeEntry();
  }

}
