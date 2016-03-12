package enterprises.orbital.evekit.snapshot.common;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.common.Asset;
import enterprises.orbital.evekit.snapshot.SheetUtils;
import enterprises.orbital.evekit.snapshot.SheetUtils.DumpCell;

public class AssetSheetWriter {

  // Singleton
  private AssetSheetWriter() {}

  public static void dumpToSheet(
                                 SynchronizedEveAccount acct,
                                 ZipOutputStream stream,
                                 long at) throws IOException {
    // Sections:
    // Assets.csv
    // AssetsMeta.csv
    stream.putNextEntry(new ZipEntry("Assets.csv"));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    output.printRecord("ID", "Item ID", "Location ID", "Type ID", "Quantity", "Flag", "Singleton", "Raw Quantity", "Container");
    List<Asset> assets = new ArrayList<Asset>();
    long contid = -1;
    List<Asset> batch = Asset.getAllAssets(acct, at, 1000, contid);
    while (batch.size() > 0) {
      assets.addAll(batch);
      contid = batch.get(batch.size() - 1).getItemID();
      batch = Asset.getAllAssets(acct, at, 1000, contid);
    }

    for (Asset next : assets) {
      // @formatter:off
      SheetUtils.populateNextRow(output, 
                                 new DumpCell(next.getCid(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getItemID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getLocationID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getTypeID(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.getFlag(), SheetUtils.CellFormat.LONG_NUMBER_STYLE), 
                                 new DumpCell(next.isSingleton(), SheetUtils.CellFormat.NO_STYLE), 
                                 new DumpCell(next.getRawQuantity(), SheetUtils.CellFormat.LONG_NUMBER_STYLE),
                                 new DumpCell(next.getContainer(), SheetUtils.CellFormat.NO_STYLE)); 
      // @formatter:on
    }
    output.flush();
    stream.closeEntry();

    // Handle MetaData
    output = SheetUtils.prepForMetaData("AssetsMeta.csv", stream, false, null);
    for (Asset next : assets) {
      int count = SheetUtils.dumpNextMetaData(acct, output, next.getCid(), "Asset");
      if (count > 0) output.println();
    }
    output.flush();
    stream.closeEntry();
  }
}
