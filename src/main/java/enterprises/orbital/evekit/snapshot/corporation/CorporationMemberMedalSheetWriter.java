package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.CorporationMemberMedal;
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

public class CorporationMemberMedalSheetWriter {

  // Singleton
  private CorporationMemberMedalSheetWriter() {}

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // MemberMedals.csv
    // MemberMedalsMeta.csv
    stream.putNextEntry(new ZipEntry("MemberMedals.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Medal ID", "Character ID", "Issued (Raw)", "Issued", "Issuer ID", "Reason", "Status");
    List<Long> metaIDs = new ArrayList<>();
    List<CorporationMemberMedal> batch = CachedData.retrieveAll(at,
                                                                (contid, at1) -> CorporationMemberMedal.accessQuery(
                                                                    acct, contid, 1000, false,
                                                                    at1,
                                                                    AttributeSelector.any(), AttributeSelector.any(),
                                                                    AttributeSelector.any(), AttributeSelector.any(),
                                                                    AttributeSelector.any(), AttributeSelector.any()));

    for (CorporationMemberMedal next : batch) {
      // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getMedalID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getIssued(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getIssued()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getIssuerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getReason(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getStatus(), SheetUtils.CellFormat.NO_STYLE)); 
        // @formatter:on
      metaIDs.add(next.getCid());
    }

    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("MemberMedalsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CorporationMemberMedal");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
