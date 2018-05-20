package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.*;
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

public class FleetsSheetWriter {

  // Singleton
  private FleetsSheetWriter() {}

  private static List<Long> dumpCharacterFleets(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("CharacterFleets.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Fleet ID", "Role", "Squad ID", "Wing ID");

    List<CharacterFleet> allFleets = CachedData.retrieveAll(at,
                                                            (contid, at1) -> CharacterFleet.accessQuery(acct, contid,
                                                                                                        1000,
                                                                                                        false,
                                                                                                        at1,
                                                                                                        AttributeSelector.any(),
                                                                                                        AttributeSelector.any(),
                                                                                                        AttributeSelector.any(),
                                                                                                        AttributeSelector.any()));
    for (CharacterFleet next : allFleets) {
      // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getFleetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getRole(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getSquadID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getWingID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
        // @formatter:on
      if (next.hasMetaData()) itemIDs.add(next.getCid());
    }
    if (allFleets.size() > 0) output.println();
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  private static List<Long> dumpFleetInfo(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("FleetInfo.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Fleet ID", "Is free move?", "Is registered?", "Is voice enabled?", "MOTD");

    List<FleetInfo> allFleets = CachedData.retrieveAll(at,
                                                       (contid, at1) -> FleetInfo.accessQuery(acct, contid,
                                                                                              1000,
                                                                                              false,
                                                                                              at1,
                                                                                              AttributeSelector.any(),
                                                                                              AttributeSelector.any(),
                                                                                              AttributeSelector.any(),
                                                                                              AttributeSelector.any(),
                                                                                              AttributeSelector.any()));
    for (FleetInfo next : allFleets) {
      // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getFleetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.isFreeMove(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.isRegistered(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.isVoiceEnabled(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getMotd(), SheetUtils.CellFormat.NO_STYLE));
        // @formatter:on
      if (next.hasMetaData()) itemIDs.add(next.getCid());
    }
    if (allFleets.size() > 0) output.println();
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  private static List<Long> dumpFleetMembers(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("FleetMembers.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Fleet ID", "Character ID", "Join Time (Raw)", "Join Time", "Role",
                       "Role Name", "Ship Type ID", "Solar System ID", "Squad ID", "Station ID", "Takes fleet warp?",
                       "Wing ID");

    List<FleetMember> allMembers = CachedData.retrieveAll(at,
                                                          (contid, at1) -> FleetMember.accessQuery(acct, contid, 1000,
                                                                                                   false,
                                                                                                   at1,
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
    for (FleetMember next : allMembers) {
      // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getFleetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getJoinTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getJoinTime()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getRole(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getRoleName(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getSquadID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getStationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.isTakesFleetWarp(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getWingID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
        // @formatter:on
      if (next.hasMetaData()) itemIDs.add(next.getCid());
    }
    if (allMembers.size() > 0) output.println();
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  private static List<Long> dumpFleetWings(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("FleetWings.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Fleet ID", "Wing ID", "Name");

    List<FleetWing> allWings = CachedData.retrieveAll(at,
                                                      (contid, at1) -> FleetWing.accessQuery(acct, contid, 1000,
                                                                                             false,
                                                                                             at1,
                                                                                             AttributeSelector.any(),
                                                                                             AttributeSelector.any(),
                                                                                             AttributeSelector.any()));
    for (FleetWing next : allWings) {
      // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getFleetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getWingID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getName(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
        // @formatter:on
      if (next.hasMetaData()) itemIDs.add(next.getCid());
    }
    if (allWings.size() > 0) output.println();
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  private static List<Long> dumpFleetSquads(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("FleetSquads.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Fleet ID", "Wing ID", "Squad ID", "Name");

    List<FleetSquad> allSquads = CachedData.retrieveAll(at,
                                                        (contid, at1) -> FleetSquad.accessQuery(acct, contid, 1000,
                                                                                                false,
                                                                                                at1,
                                                                                                AttributeSelector.any(),
                                                                                                AttributeSelector.any(),
                                                                                                AttributeSelector.any(),
                                                                                                AttributeSelector.any()));
    for (FleetSquad next : allSquads) {
      // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getFleetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getWingID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getSquadID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getName(), SheetUtils.CellFormat.NO_STYLE));
        // @formatter:on
      if (next.hasMetaData()) itemIDs.add(next.getCid());
    }
    if (allSquads.size() > 0) output.println();
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  private static void writeMetaData(List<Long> metaIDs, SynchronizedEveAccount acct, String fileName,
                                    String tableName, ZipOutputStream stream) throws IOException {
    if (metaIDs.size() > 0) {
      // Only write out meta-data if a pin was actually written.
      CSVPrinter output = SheetUtils.prepForMetaData(fileName, stream, false, null);
      for (Long next : metaIDs) {
        int count = SheetUtils.dumpNextMetaData(acct, output, next, tableName);
        if (count > 0) output.println();
      }
      output.flush();
      stream.closeEntry();
    }
  }

  public static void dumpToSheet(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    // Sections:
    // CharacterFleets.csv
    // CharacterFleetsMeta.csv
    // FleetInfo.csv
    // FleetInfoMeta.csv
    // FleetMembers.csv
    // FleetMembersMeta.csv
    // FleetWings.csv
    // FleetWingsMeta.csv
    // FleetSquads.csv
    // FleetSquadsMeta.csv
    List<Long> metaIDs;

    // Write out character fleets
    metaIDs = dumpCharacterFleets(acct, stream, at);
    writeMetaData(metaIDs, acct, "CharacterFleetsMeta.csv", "CharacterFleet", stream);

    // Write out fleet info
    metaIDs = dumpFleetInfo(acct, stream, at);
    writeMetaData(metaIDs, acct, "FleetInfoMeta.csv", "FleetInfo", stream);

    // Write out fleet members
    metaIDs = dumpFleetMembers(acct, stream, at);
    writeMetaData(metaIDs, acct, "FleetMembersMeta.csv", "FleetMember", stream);

    // Write out fleet wings
    metaIDs = dumpFleetWings(acct, stream, at);
    writeMetaData(metaIDs, acct, "FleetWingsMeta.csv", "FleetWing", stream);

    // Write out fleet squads
    metaIDs = dumpFleetSquads(acct, stream, at);
    writeMetaData(metaIDs, acct, "FleetSquadsMeta.csv", "FleetSquad", stream);
  }

}
