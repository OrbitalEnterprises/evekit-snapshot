package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.CorporationTitleRole;
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

public class CorporationTitleRoleSheetWriter {

  // Singleton
  private CorporationTitleRoleSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // CorporationTitleRoles.csv
    // CorporationTitleRolesMeta.csv
    stream.putNextEntry(new ZipEntry("CorporationTitleRoles.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Title ID", "Role Name", "Grantable?", "At HQ?", "At Base?", "At Other?");
    List<Long> metaIDs = new ArrayList<>();
    List<CorporationTitleRole> batch = CachedData.retrieveAll(at,
                                                              (contid, at1) -> CorporationTitleRole.accessQuery(acct,
                                                                                                                contid,
                                                                                                                1000,
                                                                                                                false,
                                                                                                                at1,
                                                                                                                AttributeSelector.any(),
                                                                                                                AttributeSelector.any(),
                                                                                                                AttributeSelector.any(),
                                                                                                                AttributeSelector.any(),
                                                                                                                AttributeSelector.any(),
                                                                                                                AttributeSelector.any()));

    for (CorporationTitleRole next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getTitleID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
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
    output = SheetUtils.prepForMetaData("CorporationTitleRolesMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CorporationTitleRole");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
