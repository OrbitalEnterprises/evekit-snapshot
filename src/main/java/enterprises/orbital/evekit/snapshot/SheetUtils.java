package enterprises.orbital.evekit.snapshot;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Comparator;
import java.util.Date;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TimeZone;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.CachedData;

public class SheetUtils {

  private static interface DataFormatter {
    public String format(
                         Object value);
  }

  public static enum CellFormat {
                                 NO_STYLE(new DataFormatter() {
                                   @Override
                                   public String format(
                                                        Object value) {
                                     return String.format("%s", value);
                                   }
                                 }),
                                 LONG_NUMBER_STYLE(new DataFormatter() {
                                   @Override
                                   public String format(
                                                        Object value) {
                                     return String.format("%d", value);
                                   }
                                 }),
                                 BIG_DECIMAL_STYLE(new DataFormatter() {
                                   @Override
                                   public String format(
                                                        Object value) {
                                     BigDecimal convert = ((BigDecimal) value).setScale(2);
                                     return NO_STYLE.format(convert.toPlainString());
                                   }
                                 }),
                                 DATE_STYLE(new DataFormatter() {
                                   final SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss z");

                                   @Override
                                   public String format(
                                                        Object value) {
                                     formatter.setTimeZone(TimeZone.getTimeZone("UTC"));
                                     Date convert = (Date) value;
                                     return NO_STYLE.format(formatter.format(convert));
                                   }
                                 }),
                                 DOUBLE_STYLE(new DataFormatter() {
                                   @Override
                                   public String format(
                                                        Object value) {
                                     return String.format("%f", value);
                                   }
                                 });

    private DataFormatter formatType;

    private CellFormat(DataFormatter type) {
      formatType = type;
    }

    public String format(
                         Object val) {
      return formatType.format(val);
    }
  }

  // Singletone
  private SheetUtils() {}

  public static class DumpCell {
    public Object     value;
    public CellFormat format;

    public DumpCell(Object v, CellFormat f) {
      value = v;
      format = f;
    }
  }

  public static final Comparator<CachedData> ascendingCachedDataComparator = new Comparator<CachedData>() {

    @Override
    public int compare(
                       CachedData o1,
                       CachedData o2) {
      long o1Cid = o1.getCid();
      long o2Cid = o2.getCid();
      if (o1Cid < o2Cid) return -1;
      if (o1Cid == o2Cid) return 0;
      return 1;
    }

  };

  public static void populateNextRow(
                                     CSVPrinter output,
                                     DumpCell... cells) throws IOException {
    for (DumpCell next : cells) {
      output.print(next.format.format(next.value));
    }
    output.println();
  }

  public static CSVPrinter prepForMetaData(
                                           String file,
                                           ZipOutputStream stream,
                                           boolean skipHeader,
                                           String alternateTitle) throws IOException {
    stream.putNextEntry(new ZipEntry(file));
    CSVPrinter output = CSVFormat.EXCEL.print(new OutputStreamWriter(stream));
    if (!skipHeader) {
      output.printRecord(alternateTitle != null ? alternateTitle : "Meta Data");
      output.printRecord("ID", "Key", "Value");
    }
    return output;
  }

  public static int dumpNextMetaData(
                                     SynchronizedEveAccount acct,
                                     CSVPrinter output,
                                     long metaID,
                                     String tableName) throws IOException {
    CachedData tagged = CachedData.get(metaID, tableName);
    if (tagged == null) return 0;
    Set<Entry<String, String>> allMD = tagged.getAllMetaData();
    for (Entry<String, String> next : allMD) {
      // @formatter:off
      SheetUtils.populateNextRow(output,
                                 new DumpCell(metaID, CellFormat.NO_STYLE),
                                 new DumpCell(next.getKey(), CellFormat.NO_STYLE),
                                 new DumpCell(next.getValue(), CellFormat.NO_STYLE));
      // @formatter:on
    }

    return allMD.size();
  }
}
