package enterprises.orbital.evekit.snapshot.corporation;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.corporation.Fuel;
import enterprises.orbital.evekit.model.corporation.Starbase;
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

public class StarbaseSheetWriter {

  // Singleton
  private StarbaseSheetWriter() {}

  private static List<Long> dumpFuel(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      List<Long> stationIDs,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("StarbaseFuel.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Starbase ID", "Type ID", "Quantity");
    for (long stationID : stationIDs) {
      List<Fuel> batch = CachedData.retrieveAll(at, (contid, at1) -> Fuel.accessQuery(acct, contid, 1000, false, at1,
                                                                                      AttributeSelector.values(
                                                                                          stationID),
                                                                                      AttributeSelector.any(),
                                                                                      AttributeSelector.any()));
      if (batch.size() > 0) {
        for (Fuel next : batch) {
          // @formatter:off
          SheetUtils.populateNextRow(output,
                                     new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getStarbaseID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
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
    // StarbaseFuel.csv
    // StarbaseFuelMeta.csv
    stream.putNextEntry(new ZipEntry("Starbases.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Starbase ID", "Type ID", "System ID", "Moon ID", "State", "Unanchor At (Raw)",
                       "Unanchor At", "Reinforced Until (Raw)", "Reinforced Until", "Onlined Since (Raw)",
                       "Onlined Since", "Fuel Bay View", "Fuel Bay Take", "Anchor", "Unanchor", "Online",
                       "Offline", "Allow Corporation Members", "Allow Alliance Members",
                       "Use Alliance Standings", "Attack Standing Threshold", "Attack Security Status Threshold",
                       "Attack If Other Security Status Dropping", "Attack If At War");
    List<Starbase> batch = CachedData.retrieveAll(at,
                                                  (contid, at1) -> Starbase.accessQuery(acct, contid, 1000, false, at1,
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any(),
                                                                                        AttributeSelector.any()));
    List<Long> starbaseIDs = new ArrayList<>();

    // Write out starbase data first
    for (Starbase next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getStarbaseID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getMoonID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getState(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getUnanchorAt(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getUnanchorAt()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getReinforcedUntil(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getReinforcedUntil()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getOnlinedSince(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(new Date(next.getOnlinedSince()), SheetUtils.CellFormat.DATE_STYLE),
                                 new DumpCell(next.getFuelBayView(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getFuelBayTake(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getAnchor(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getUnanchor(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getOnline(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getOffline(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.isAllowCorporationMembers(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.isAllowAllianceMembers(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.isUseAllianceStandings(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getAttackStandingThreshold(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(next.getAttackSecurityStatusThreshold(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                 new DumpCell(next.isAttackIfOtherSecurityStatusDropping(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.isAttackIfAtWar(), SheetUtils.CellFormat.NO_STYLE));
      // @formatter:on
      starbaseIDs.add(next.getStarbaseID());
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

      // Write out fuel data in the same style as meta data.
      List<Long> writes = dumpFuel(acct, stream, starbaseIDs, at);
      if (writes.size() > 0) {
        // Only write out meta-data if fuel was actually written.
        output = SheetUtils.prepForMetaData("StarbaseFuelMeta.csv", stream, false, null);
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
