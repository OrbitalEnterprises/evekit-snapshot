package enterprises.orbital.evekit.snapshot.common;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.common.Kill;
import enterprises.orbital.evekit.model.common.KillAttacker;
import enterprises.orbital.evekit.model.common.KillItem;
import enterprises.orbital.evekit.model.common.KillVictim;
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

public class KillSheetWriter {

  // Singleton
  private KillSheetWriter() {}

  private static List<Long> dumpKillAttackers(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    stream.putNextEntry(new ZipEntry("KillAttackers.csv"));
    List<Long> metaIDs = new ArrayList<>();

    try (CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream))) {
      output.printRecord("ID", "Kill ID", "Attacker Character ID", "Alliance ID", "Attacker Corporation ID",
                         "Damage Done", "Faction ID", "Security Status", "Ship Type ID", "Weapon Type ID",
                         "Final Blow");

      CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
      CachedData.stream(at, (long contid, AttributeSelector ats) ->
                            KillAttacker.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                                     AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                     AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                     AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()),
                        true, capture)
                .forEach(next -> {
                  try {
                    // @formatter:off
                  SheetUtils.populateNextRow(output,
                                             new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getKillID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getAttackerCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getAttackerCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getDamageDone(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getFactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getSecurityStatus(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                             new DumpCell(next.getShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getWeaponTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.isFinalBlow(), SheetUtils.CellFormat.NO_STYLE));
                  // @formatter:on
                  } catch (IOException e) {
                    capture.handle(e);
                  }
                  metaIDs.add(next.getCid());
                });

      output.flush();
    }

    return metaIDs;
  }

  private static List<Long> dumpKillItems(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    List<Long> metaIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("KillItems.csv"));

    try (CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream))) {
      output.printRecord("ID", "Kill ID", "Type ID", "Flag", "Quantity Destroyed", "Quantity Dropped", "Singleton",
                         "Container", "Sequence");

      CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
      CachedData.stream(at, (long contid, AttributeSelector ats) ->
                            KillItem.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                                 AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                 AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                 AttributeSelector.any()),
                        true, capture)
                .forEach(next -> {
                  try {
                    // @formatter:off
                  SheetUtils.populateNextRow(output,
                                             new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getKillID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getFlag(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getQtyDestroyed(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getQtyDropped(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getSingleton(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getContainerSequence(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getSequence(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
                  // @formatter:on
                  } catch (IOException e) {
                    capture.handle(e);
                  }
                  metaIDs.add(next.getCid());
                });

      output.flush();
    }

    return metaIDs;
  }

  private static List<Long> dumpKillVictims(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      long at) throws IOException {
    stream.putNextEntry(new ZipEntry("KillVictims.csv"));
    List<Long> metaIDs = new ArrayList<>();

    try (CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream))) {
      output.printRecord("ID", "Kill ID", "Alliance ID", "Kill Character ID", "Kill Corporation ID",
                         "Damage Taken", "Faction ID", "Faction Name", "Ship Type ID", "X", "Y", "Z");

      CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
      CachedData.stream(at, (long contid, AttributeSelector ats) ->
                            KillVictim.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                                   AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                   AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                                   AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any()),
                        true, capture)
                .forEach(next -> {
                  try {
                    // @formatter:off
                  SheetUtils.populateNextRow(output,
                                             new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getKillID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getKillCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getKillCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getDamageTaken(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getFactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
                  // @formatter:on
                  } catch (IOException e) {
                    capture.handle(e);
                  }
                  metaIDs.add(next.getCid());
                });

      output.flush();
    }

    return metaIDs;
  }

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Kills.csv
    // KillsMeta.csv
    // KillAttackers.csv
    // KillAttackersMeta.csv
    // KillItems.csv
    // KillItemsMeta.csv
    // KillVictims.csv
    // KillVictimsMeta.csv
    stream.putNextEntry(new ZipEntry("Kills.csv"));
    final List<Long> metaIDs = new ArrayList<>();

    try (CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream))) {
      output.printRecord("ID", "Kill ID", "Kill Time (Raw)", "Kill Time", "Moon ID", "Solar System ID");

      CachedData.SimpleStreamExceptionHandler capture = new CachedData.SimpleStreamExceptionHandler();
      CachedData.stream(at, (long contid, AttributeSelector ats) ->
                            Kill.accessQuery(acct, contid, 1000, false, ats, AttributeSelector.any(),
                                             AttributeSelector.any(), AttributeSelector.any(), AttributeSelector.any(),
                                             AttributeSelector.any()),
                        true, capture)
                .forEach(next -> {
                  try {
                    // @formatter:off
                  SheetUtils.populateNextRow(output,
                                             new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                             new DumpCell(next.getKillID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getKillTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(new Date(next.getKillTime()), SheetUtils.CellFormat.DATE_STYLE),
                                             new DumpCell(next.getMoonID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                             new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
                  // @formatter:on
                  } catch (IOException e) {
                    capture.handle(e);
                  }
                  metaIDs.add(next.getCid());
                });

      output.flush();
    }

    if (!metaIDs.isEmpty()) {
      // Wrote at least one kill so proceed

      // Write out meta data for kills
      try (CSVPrinter output = SheetUtils.prepForMetaData("KillsMeta.csv", stream, false, null)) {
        for (Long next : metaIDs) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "Kill");
          if (count > 0) output.println();
        }
        output.flush();
      }
      stream.closeEntry();

      // Write out kill attackers data in the same style as meta data.
      metaIDs.clear();
      metaIDs.addAll(dumpKillAttackers(acct, stream, at));
      if (!metaIDs.isEmpty()) {
        // Only write out meta-data if an attacker was actually written.
        try (CSVPrinter output = SheetUtils.prepForMetaData("KillAttackersMeta.csv", stream, false, null)) {
          for (Long next : metaIDs) {
            int count = SheetUtils.dumpNextMetaData(acct, output, next, "KillAttacker");
            if (count > 0) output.println();
          }
          output.flush();
        }
      }

      // Write out kill items data in the same style as meta data.
      metaIDs.clear();
      metaIDs.addAll(dumpKillItems(acct, stream, at));
      if (!metaIDs.isEmpty()) {
        // Only write out meta-data if an item was actually written.
        try (CSVPrinter output = SheetUtils.prepForMetaData("KillItemsMeta.csv", stream, false, null)) {
          for (Long next : metaIDs) {
            int count = SheetUtils.dumpNextMetaData(acct, output, next, "KillItem");
            if (count > 0) output.println();
          }
          output.flush();
        }
      }

      // Write out kill victim data in the same style as meta data.
      metaIDs.clear();
      metaIDs.addAll(dumpKillVictims(acct, stream, at));
      if (!metaIDs.isEmpty()) {
        // Only write out meta-data if a victim was actually written.
        try (CSVPrinter output = SheetUtils.prepForMetaData("KillVictimsMeta.csv", stream, false, null)) {
          for (Long next : metaIDs) {
            int count = SheetUtils.dumpNextMetaData(acct, output, next, "KillVictim");
            if (count > 0) output.println();
          }
          output.flush();
        }
      }
    }
  }

}
