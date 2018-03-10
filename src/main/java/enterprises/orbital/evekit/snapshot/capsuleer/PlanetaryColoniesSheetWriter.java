package enterprises.orbital.evekit.snapshot.capsuleer;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import enterprises.orbital.evekit.model.character.PlanetaryColony;
import enterprises.orbital.evekit.model.character.PlanetaryLink;
import enterprises.orbital.evekit.model.character.PlanetaryPin;
import enterprises.orbital.evekit.model.character.PlanetaryRoute;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class PlanetaryColoniesSheetWriter {

  // Singleton
  private PlanetaryColoniesSheetWriter() {}

  private static List<Long> dumpPlanetaryPins(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      List<Integer> planets,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("PlanetaryPins.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Planet ID", "Pin ID", "Type ID", "Schematic ID", "Last Cycle Start (Raw)",
                       "Last Cycle Start", "Cycle Time",
                       "Quantity Per Cycle", "Install Time (Raw)", "Install Time", "Expiry Time (Raw)", "Expiry Time",
                       "Product Type ID",
                       "Longitude", "Latitude", "Head Radius", "Heads", "Contents");
    for (int planetID : planets) {
      List<PlanetaryPin> allPins = CachedData.retrieveAll(at,
                                                          (contid, at1) -> PlanetaryPin.accessQuery(acct, contid, 1000,
                                                                                                    false,
                                                                                                    at1,
                                                                                                    AttributeSelector.values(
                                                                                                        planetID),
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
      for (PlanetaryPin next : allPins) {
        // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getPlanetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getPinID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getSchematicID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getLastCycleStart(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getLastCycleStart()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getCycleTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getQuantityPerCycle(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getInstallTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getInstallTime()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getExpiryTime(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(new Date(next.getExpiryTime()), SheetUtils.CellFormat.DATE_STYLE),
                                   new DumpCell(next.getProductTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getLongitude(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                   new DumpCell(next.getLatitude(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                   new DumpCell(next.getHeadRadius(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                   new DumpCell(Arrays.toString(next.getHeads()
                                                                    .stream()
                                                                    .map(x -> "[" + x.getHeadID() + ", " + x.getLongitude() + ", " + x.getLatitude() + "]")
                                                                    .toArray(String[]::new)), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(Arrays.toString(next.getContents()
                                                                    .stream()
                                                                    .map(x -> "[" + x.getTypeID() + ", " + x.getAmount() + "]")
                                                                    .toArray(String[]::new)), SheetUtils.CellFormat.NO_STYLE));
        // @formatter:on
        if (next.hasMetaData()) itemIDs.add(next.getCid());
      }
      if (allPins.size() > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  private static List<Long> dumpPlanetaryLinks(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      List<Integer> planets,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("PlanetaryLinks.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Planet ID", "Source Pin ID", "Destination Pin ID", "Link Level");
    for (int planetID : planets) {
      List<PlanetaryLink> allLinks = CachedData.retrieveAll(at, (contid, at1) -> PlanetaryLink.accessQuery(acct, contid,
                                                                                                           1000, false,
                                                                                                           at1,
                                                                                                           AttributeSelector.values(
                                                                                                               planetID),
                                                                                                           AttributeSelector.any(),
                                                                                                           AttributeSelector.any(),
                                                                                                           AttributeSelector.any()));
      for (PlanetaryLink next : allLinks) {
        // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getPlanetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getSourcePinID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getDestinationPinID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getLinkLevel(), SheetUtils.CellFormat.LONG_NUMBER_STYLE));
        // @formatter:on
        if (next.hasMetaData()) itemIDs.add(next.getCid());
      }
      if (allLinks.size() > 0) output.println();
    }
    output.flush();
    stream.closeEntry();

    return itemIDs;
  }

  private static List<Long> dumpPlanetaryRoutes(
      SynchronizedEveAccount acct,
      ZipOutputStream stream,
      List<Integer> planets,
      long at) throws IOException {
    List<Long> itemIDs = new ArrayList<>();
    stream.putNextEntry(new ZipEntry("PlanetaryRoutes.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Planet ID", "Route ID", "Source Pin ID", "Destination Pin ID", "Content Type ID",
                       "Quantity", "Waypoints");
    for (int planetID : planets) {
      List<PlanetaryRoute> allRoutes = CachedData.retrieveAll(at,
                                                              (contid, at1) -> PlanetaryRoute.accessQuery(acct, contid,
                                                                                                          1000, false,
                                                                                                          at1,
                                                                                                          AttributeSelector.values(
                                                                                                              planetID),
                                                                                                          AttributeSelector.any(),
                                                                                                          AttributeSelector.any(),
                                                                                                          AttributeSelector.any(),
                                                                                                          AttributeSelector.any(),
                                                                                                          AttributeSelector.any(),
                                                                                                          AttributeSelector.any()));
      for (PlanetaryRoute next : allRoutes) {
        // @formatter:off
        SheetUtils.populateNextRow(output,
                                   new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE),
                                   new DumpCell(next.getPlanetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getRouteID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getSourcePinID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getDestinationPinID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getContentTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                   new DumpCell(next.getQuantity(), SheetUtils.CellFormat.DOUBLE_STYLE),
                                   new DumpCell(Arrays.toString(next.getWaypoints().toArray(new Long[next.getWaypoints().size()])), SheetUtils.CellFormat.NO_STYLE));
        // @formatter:on
        if (next.hasMetaData()) itemIDs.add(next.getCid());
      }
      if (allRoutes.size() > 0) output.println();
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
    // PlanetaryColonies.csv
    // PlanetaryColoniesMeta.csv
    // PlanetaryPins.csv
    // PlanetaryPinsMeta.csv
    // PlanetaryLinks.csv
    // PlanetaryLinksMeta.csv
    // PlanetaryRoutes.csv
    // PlanetaryRoutesMeta.csv
    stream.putNextEntry(new ZipEntry("PlanetaryColonies.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Planet ID", "Solar System ID", "Planet Type", "Owner ID",
                       "Last Update (Raw)", "Last Update", "Upgrade Level", "Number Of Pins");
    List<Integer> planetIDs = new ArrayList<>();
    List<Long> metaIDs = new ArrayList<>();
    List<PlanetaryColony> batch = CachedData.retrieveAll(at,
                                                         (contid, at1) -> PlanetaryColony.accessQuery(acct, contid,
                                                                                                      1000, false,
                                                                                                      at1,
                                                                                                      AttributeSelector.any(),
                                                                                                      AttributeSelector.any(),
                                                                                                      AttributeSelector.any(),
                                                                                                      AttributeSelector.any(),
                                                                                                      AttributeSelector.any(),
                                                                                                      AttributeSelector.any(),
                                                                                                      AttributeSelector.any()));

    for (PlanetaryColony next : batch) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getPlanetID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getSolarSystemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getPlanetType(), SheetUtils.CellFormat.NO_STYLE),
                                 new DumpCell(next.getOwnerID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getLastUpdate(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(new Date(next.getLastUpdate()), SheetUtils.CellFormat.DATE_STYLE), 
                                 new DumpCell(next.getUpgradeLevel(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getNumberOfPins(), SheetUtils.CellFormat.LONG_NUMBER_STYLE)); 
      // @formatter:on
      planetIDs.add(next.getPlanetID());
      if (next.hasMetaData()) metaIDs.add(next.getCid());
    }

    output.flush();
    stream.closeEntry();

    if (planetIDs.size() > 0) {
      // Wrote at least one colony so proceed

      // Write out meta data for colonies
      output = SheetUtils.prepForMetaData("PlanetaryColoniesMeta.csv", stream, false, null);
      for (Long next : metaIDs) {
        int count = SheetUtils.dumpNextMetaData(acct, output, next, "PlanetaryColony");
        if (count > 0) output.println();
      }
      output.flush();
      stream.closeEntry();

      // Write out planetary pins data in the same style as meta data.
      metaIDs = dumpPlanetaryPins(acct, stream, planetIDs, at);
      if (metaIDs.size() > 0) {
        // Only write out meta-data if a pin was actually written.
        output = SheetUtils.prepForMetaData("PlanetaryPinsMeta.csv", stream, false, null);
        for (Long next : metaIDs) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "PlanetaryPin");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }

      // Write out links data in the same style as meta data.
      metaIDs = dumpPlanetaryLinks(acct, stream, planetIDs, at);
      if (metaIDs.size() > 0) {
        // Only write out meta-data if a link was actually written.
        output = SheetUtils.prepForMetaData("PlanetaryLinksMeta.csv", stream, false, null);
        for (Long next : metaIDs) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "PlanetaryLink");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }

      // Write out routes data in the same style as meta data.
      metaIDs = dumpPlanetaryRoutes(acct, stream, planetIDs, at);
      if (metaIDs.size() > 0) {
        // Only write out meta-data if a route was actually written.
        output = SheetUtils.prepForMetaData("PlanetaryRoutesMeta.csv", stream, false, null);
        for (Long next : metaIDs) {
          int count = SheetUtils.dumpNextMetaData(acct, output, next, "PlanetaryRoute");
          if (count > 0) output.println();
        }
        output.flush();
        stream.closeEntry();
      }
    }
  }

}
