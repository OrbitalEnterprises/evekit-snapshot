package enterprises.orbital.evekit.snapshot.capsuleer;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.character.CharacterMedal;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class MedalSheetWriter {

  // Singleton
  private MedalSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Medals.csv
    // MedalsMeta.csv
    stream.putNextEntry(new ZipEntry("Medals.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Medal ID", "Title", "Description", "Corporation ID", "Issued (Raw)", "Issued", "Issuer ID", "Reason", "Status");
    List<Long> metaIDs = new ArrayList<Long>();
    List<CharacterMedal> medals = CharacterMedal.getAllMedals(acct, at);
    for (CharacterMedal next : medals) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getMedalID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTitle(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getDescription(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
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
    output = SheetUtils.prepForMetaData("MedalsMeta.csv", stream, false, null);
    for (Long next : metaIDs) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next, "CharacterMedal");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }

}
