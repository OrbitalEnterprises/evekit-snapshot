package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.MemberTracking;
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

public class MemberTrackingSheetWriter {

  // Singleton
  private MemberTrackingSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // MemberTracking.csv
    // MemberTrackingMeta.csv
    stream.putNextEntry(new ZipEntry("MemberTracking.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Character ID", "Base ID", "Location ID", "Logoff Date Time (Raw)", "Logoff Date Time",
                       "Logon Date Time (Raw)", "Logon Date Time", "Ship Type ID", "Start Date Time (Raw)",
                       "Start Date Time");
    List<Long> metaIDs = new ArrayList<>();
    List<MemberTracking> batch = CachedData.retrieveAll(at,
                                                        (contid, at1) -> MemberTracking.accessQuery(acct, contid, 1000,
                                                                                                    false,
                                                                                                    at1,
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any(),
                                                                                                    AttributeSelector.any()));

    for (MemberTracking next : batch) {
      // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getBaseID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getLogoffDateTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getLogoffDateTime()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getLogonDateTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getLogonDateTime()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getStartDateTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getStartDateTime()), SheetUtils.CellFormat.DATE_STYLE));
        // @formatter:on
      metaIDs.add(next.getCid());
    }

    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MemberTrackingMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "MemberTracking");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
