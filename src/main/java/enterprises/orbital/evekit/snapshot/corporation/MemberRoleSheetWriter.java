package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.MemberRole;
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

public class MemberRoleSheetWriter {

  // Singleton
  private MemberRoleSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // MemberRoles.csv
    // MemberRolesMeta.csv
    stream.putNextEntry(new ZipEntry("MemberRoles.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Character ID", "Role Name", "Grantable?", "At HQ?", "At Base?", "At Other?");
    List<Long> metaIDs = new ArrayList<>();
    List<MemberRole> batch = CachedData.retrieveAll(at,
                                                    (contid, at1) -> MemberRole.accessQuery(acct, contid, 1000, false,
                                                                                            at1,
                                                                                            AttributeSelector.any(),
                                                                                            AttributeSelector.any(),
                                                                                            AttributeSelector.any(),
                                                                                            AttributeSelector.any(),
                                                                                            AttributeSelector.any(),
                                                                                            AttributeSelector.any()));


    for (MemberRole next : batch) {
      // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getRoleName(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.isGrantable(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.isAtHQ(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.isAtBase(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.isAtOther(), SheetUtils.CellFormat.NO_STYLE));
        // @formatter:on
      metaIDs.add(next.getCid());
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MemberRolesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "MemberRole");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
