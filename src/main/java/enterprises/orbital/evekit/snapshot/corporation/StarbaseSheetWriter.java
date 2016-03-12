package enterprises.orbital.evekit.snapshot.corporation;

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
import enterprises.orbital.evekit.model.corporation.Fuel;
import enterprises.orbital.evekit.model.corporation.Starbase;
import enterprises.orbital.evekit.model.corporation.StarbaseDetail;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class StarbaseSheetWriter {

  // Singleton
  private StarbaseSheetWriter() {}

  public static List<Long> dumpStarbaseDetails(
                                               SynchronizedEveAccount acct,
                                               ZipOutputStream stream,
                                               List<Long> starbaseIDs,
                                               long at) throws IOException {
    List<Long> itemIDs = new ArrayList<Long>();
    stream.putNextEntry(new ZipEntry("StarbaseDetails.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Item ID", "State", "State Timestamp (Raw)", "State Timestamp", "Online Timestamp (Raw)", "Online Timestamp", "Usage Flags",
                       "Deploy Flags", "Allow Alliance Members", "Allow Corporation Members", "Use Standings From", "On Aggression Enabled",
                       "On Aggression Standing", "On Corporation War Enabled", "On Corporation War Standing", "On Standing Drop Enabled",
                       "On Standing Drop Standing", "On Status Drop Enabled", "On Status Drop Standing");
    for (long starbaseID : starbaseIDs) {
      StarbaseDetail next = StarbaseDetail.get(acct, at, starbaseID);
      if (next != null) {
        // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getState(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getStateTimestamp(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getStateTimestamp()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getOnlineTimestamp(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getOnlineTimestamp()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getUsageFlags(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getDeployFlags(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.isAllowAllianceMembers(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.isAllowCorporationMembers(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.isOnAggressionEnabled(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getOnAggressionStanding(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.isOnCorporationWarEnabled(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getOnCorporationWarStanding(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.isOnStandingDropEnabled(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getOnStandingDropStanding(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.isOnStatusDropEnabled(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getOnStandingDropStanding(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
        // @formatter:on
        itemIDs.add(next.getCid());
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static List<Long> dumpFuel(
                                    SynchronizedEveAccount acct,
                                    ZipOutputStream stream,
                                    List<Long> stationIDs,
                                    long at) throws IOException {
    List<Long> itemIDs = new ArrayList<Long>();
    stream.putNextEntry(new ZipEntry("StarbaseFuel.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Item ID", "Type ID", "Quantity");
    for (long stationID : stationIDs) {
      List<Fuel> batch = Fuel.getAllByItemID(acct, at, stationID);
      if (batch.size() > 0) {
        for (Fuel next : batch) {
          // @formatter:off
          SheetUtils.populateNextRow(output,
                                     new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
          // @formatter:on
          itemIDs.add(next.getCid());
        }
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Starbases.csv
    // StarbasesMeta.csv
    // StarbaseDetails.csv
    // StarbaseDetailsMeta.csv
    // Fuel.csv
    // FuelMeta.csv
    stream.putNextEntry(new ZipEntry("Starbases.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Item ID", "Location ID", "Moon ID", "Online Timestamp (Raw)", "Online Timestamp", "State", "State Timestamp (Raw)",
                       "State Timestamp", "Type ID", "Standing Owner ID");
    List<Starbase> batch = Starbase.getAll(acct, at);
    List<Long> starbaseIDs = new ArrayList<Long>();

    // Write out starbase data first
    for (Starbase next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getMoonID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getOnlineTimestamp(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(next.getOnlineTimestamp()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(next.getState(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getStateTimestamp(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getStateTimestamp()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getStandingOwnerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
      // @formatter:on
      starbaseIDs.add(next.getItemID());
    }
    output.flush();
    stream.closeEntry();

    if (batch.size() > 0) {
      // Wrote at least one starbase so proceed

      // Write out meta data for starbases
      output = SheetUtils.prepForMetaData("StarbasesMeta.csv", stream, false, null);
      for (Starbase next : batch) {
        int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Starbase");
        if (count > 0) output.println();
      }
      output.flush();
      stream.closeEntry();

      // Write out starbase datails data in the same style as meta data.
      List<Long> writes = dumpStarbaseDetails(acct, stream, starbaseIDs, at);
      if (writes.size() > 0) {
        // Only write out meta-data if a starbase detail was actually written.
        output = SheetUtils.prepForMetaData("StarbaseDetailsMeta.csv", stream, false, null);
        for (Long next : writes) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "StarbaseDetail");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }

      // Write out fuel data in the same style as meta data.
      writes = dumpFuel(acct, stream, starbaseIDs, at);
      if (writes.size() > 0) {
        // Only write out meta-data if fuel was actually written.
        output = SheetUtils.prepForMetaData("FuelMeta.csv", stream, false, null);
        for (Long next : writes) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "Fuel");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }
    }
  }

}
