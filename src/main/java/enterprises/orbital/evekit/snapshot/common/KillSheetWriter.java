package enterprises.orbital.evekit.snapshot.common;

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
import enterprises.orbital.evekit.model.common.Kill;
import enterprises.orbital.evekit.model.common.KillAttacker;
import enterprises.orbital.evekit.model.common.KillItem;
import enterprises.orbital.evekit.model.common.KillVictim;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class KillSheetWriter {

  // Singleton
  private KillSheetWriter() {}

  public static List<Long> dumpKillAttackers(
                                             SynchronizedEveAccount acct,
                                             ZipOutputStream stream,
                                             List<Long> kills,
                                             long at) throws IOException {
    List<Long> itemIDs = new ArrayList<Long>();
    stream.putNextEntry(new ZipEntry("KillAttackers.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Kill ID", "Attacker Character ID", "Alliance ID", "Alliance Name", "Attacker Character Name", "Attacker Corporation ID",
                       "Attacker Corporation Name", "Damage Done", "Faction ID", "Faction Name", "Security Status", "Ship Type ID", "Weapon Type ID",
                       "Final Blow");
    for (long killID : kills) {
      List<KillAttacker> allAttackers = new ArrayList<KillAttacker>();
      long contid = -1;
      List<KillAttacker> batch = KillAttacker.getAllKillAttackers(acct, at, killID, 1000, contid);
      while (batch.size() > 0) {
        allAttackers.addAll(batch);
        contid = batch.get(batch.size() - 1).getCid();
        batch = KillAttacker.getAllKillAttackers(acct, at, killID, 1000, contid);
      }
      if (allAttackers.size() > 0) {
        for (KillAttacker next : allAttackers) {
          // @formatter:off
          SheetUtils.populateNextRow(output,
                                     new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getKillID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getAttackerCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getAllianceName(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getAttackerCharacterName(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getAttackerCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getAttackerCorporationName(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getDamageDone(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getFactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getFactionName(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getSecurityStatus(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                     new DumpCell(next.getShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getWeaponTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.isFinalBlow(), SheetUtils.CellFormat.NO_STYLE));
          // @formatter:on
          if (next.hasMetaData()) itemIDs.add(next.getCid());
        }
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static List<Long> dumpKillItems(
                                         SynchronizedEveAccount acct,
                                         ZipOutputStream stream,
                                         List<Long> kills,
                                         long at) throws IOException {
    List<Long> itemIDs = new ArrayList<Long>();
    stream.putNextEntry(new ZipEntry("KillItems.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Kill ID", "Type ID", "Flag", "Quantity Destroyed", "Quantity Dropped", "Singleton", "Container", "Sequence");
    for (long killID : kills) {
      List<KillItem> allItems = new ArrayList<KillItem>();
      int contid = 0;
      List<KillItem> batch = KillItem.getAllKillItems(acct, at, killID, 1000, contid);
      while (batch.size() > 0) {
        allItems.addAll(batch);
        contid = batch.get(batch.size() - 1).getSequence();
        batch = KillItem.getAllKillItems(acct, at, killID, 1000, contid);
      }
      if (allItems.size() > 0) {
        for (KillItem next : allItems) {
          // @formatter:off
          SheetUtils.populateNextRow(output,
                                     new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getKillID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getFlag(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getQtyDestroyed(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getQtyDropped(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.isSingleton(), SheetUtils.CellFormat.NO_STYLE),
                                     new DumpCell(next.getContainerSequence(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                     new DumpCell(next.getSequence(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
          // @formatter:on
          if (next.hasMetaData()) itemIDs.add(next.getCid());
        }
        output.println();
      }
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  public static List<Long> dumpKillVictims(
                                           SynchronizedEveAccount acct,
                                           ZipOutputStream stream,
                                           List<Long> kills,
                                           long at) throws IOException {
    List<Long> itemIDs = new ArrayList<Long>();
    stream.putNextEntry(new ZipEntry("KillVictims.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Kill ID", "Alliance ID", "Alliance Name", "Kill Character ID", "Kill Character Name", "Kill Corporation ID",
                       "Kill Corporation Name", "Damage Taken", "Faction ID", "Faction Name", "Ship Type ID");
    for (long killID : kills) {
      KillVictim victim = KillVictim.get(acct, at, killID);
      if (victim != null) {
        // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(victim.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(victim.getKillID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(victim.getAllianceID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(victim.getAllianceName(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(victim.getKillCharacterID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(victim.getKillCharacterName(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(victim.getKillCorporationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(victim.getKillCorporationName(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(victim.getDamageTaken(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(victim.getFactionID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(victim.getFactionName(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(victim.getShipTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
        // @formatter:on
        if (victim.hasMetaData()) itemIDs.add(victim.getCid());
      }
      output.println();
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
    // Kills.csv
    // KillsMeta.csv
    // KillAttackers.csv
    // KillAttackersMeta.csv
    // KillItems.csv
    // KillItemsMeta.csv
    // KillVictims.csv
    // KillVictimsMeta.csv
    stream.putNextEntry(new ZipEntry("Kills.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Kill ID", "Kill Time (Raw)", "Kill Time", "Moon ID", "Solar System ID");
    long contid = -1;
    List<Long> killIDs = new ArrayList<Long>();
    List<Long> metaIDs = new ArrayList<Long>();
    List<Kill> batch = Kill.getKillsForward(acct, at, 1000, contid);

    while (batch.size() > 0) {
      for (Kill next : batch) {
        // @formatter:off
        SheetUtils.populateNextRow(output, 
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                   new DumpCell(next.getKillID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getKillTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(new Date(next.getKillTime()), SheetUtils.CellFormat.DATE_STYLE), 
                                   new DumpCell(next.getMoonID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                   new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
        // @formatter:on
        killIDs.add(next.getKillID());
        if (next.hasMetaData()) metaIDs.add(next.getCid());
      }

      contid = batch.get(batch.size() - 1).getKillTime();
      batch = Kill.getKillsForward(acct, at, 1000, contid);
    }
    output.flush();
    stream.closeEntry();

    if (killIDs.size() > 0) {
      // Wrote at least one kill so proceed

      // Write out meta data for kills
      output = SheetUtils.prepForMetaData("KillsMeta.csv", stream, false, null);
      for (Long next : metaIDs) {
        int count = SheetUtils.dumpNextMetaData(acct, output, next, "Kill");
        if (count > 0) output.println();
      }
      output.flush();
      stream.closeEntry();

      // Write out kill attackers data in the same style as meta data.
      metaIDs = dumpKillAttackers(acct, stream, killIDs, at);
      if (metaIDs.size() > 0) {
        // Only write out meta-data if an attacker was actually written.
        output = SheetUtils.prepForMetaData("KillAttackersMeta.csv", stream, false, null);
        for (Long next : metaIDs) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "KillAttacker");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }

      // Write out kill items data in the same style as meta data.
      metaIDs = dumpKillItems(acct, stream, killIDs, at);
      if (metaIDs.size() > 0) {
        // Only write out meta-data if an item was actually written.
        output = SheetUtils.prepForMetaData("KillItemsMeta.csv", stream, false, null);
        for (Long next : metaIDs) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "KillItem");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }

      // Write out kill victim data in the same style as meta data.
      metaIDs = dumpKillVictims(acct, stream, killIDs, at);
      if (metaIDs.size() > 0) {
        // Only write out meta-data if a victim was actually written.
        output = SheetUtils.prepForMetaData("KillVictimsMeta.csv", stream, false, null);
        for (Long next : metaIDs) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "KillVictim");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }
    }
  }

}
