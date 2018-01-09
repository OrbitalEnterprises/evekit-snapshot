package enterprises.orbital.evekit.snapshot;

import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.model.AttributeSelector;
import enterprises.orbital.evekit.model.CachedData;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class SheetUtils {

  private interface DataFormatter {
    String format(
                         Object value);
  }

  public enum CellFormat {
                                 NO_STYLE((Object value) -> String.format("%s", value)),
                                 LONG_NUMBER_STYLE((Object value) -> String.format("%d", value)),
                                 BIG_DECIMAL_STYLE(new DataFormatter() {
                                   @Override
                                   public String format(
                                                        Object value) {
                                     BigDecimal convert = ((BigDecimal) value).setScale(2);
                                     return NO_STYLE.format(convert.toPlainString());
                                   }
                                 }),
                                 DATE_STYLE(new DataFormatter() {
                                   @Override
                                   public String format(
                                                        Object value) {
                                     SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss z");
                                     formatter.setTimeZone(TimeZone.getTimeZone("UTC"));
                                     Date convert = (Date) value;
                                     return NO_STYLE.format(formatter.format(convert));
                                   }
                                 }),
                                 DOUBLE_STYLE((Object value) -> String.format("%f", value));

    private DataFormatter formatType;

    CellFormat(DataFormatter type) {
      formatType = type;
    }

    public String format(
                         Object val) {
      return formatType.format(val);
    }
  }

  // Singleton
  private SheetUtils() {}

  public static class DumpCell {
    Object     value;
    CellFormat format;

    public DumpCell(Object v, CellFormat f) {
      value = v;
      format = f;
    }
  }

  public static final Comparator<CachedData> ascendingCachedDataComparator = (o1, o2) -> {
    long o1Cid = o1.getCid();
    long o2Cid = o2.getCid();
    if (o1Cid < o2Cid) return -1;
    if (o1Cid == o2Cid) return 0;
    return 1;
  };

  public static void populateNextRow(
                                     CSVPrinter output,
                                     DumpCell... cells)
    throws IOException {
    for (DumpCell next : cells) {
      output.print(next.format.format(next.value));
    }
    output.println();
  }

  public static CSVPrinter prepForMetaData(
                                           String file,
                                           ZipOutputStream stream,
                                           boolean skipHeader,
                                           String alternateTitle)
    throws IOException {
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
                                     String tableName)
    throws IOException {
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

  public static final AttributeSelector ANY_SELECTOR = new AttributeSelector("{ any: true }");

  // Convenience function to construct a time selector for the give time.
  public static AttributeSelector makeAtSelector(long time) {
    return new AttributeSelector("{values: [" + time + "]}");
  }

  // Interface which forwards a call to the class specific query function to retrieve data
  public interface QueryCaller<A extends CachedData> {
    List<A> query(long contid, AttributeSelector at) throws IOException;
  }

  /**
   * Retrieve all data items of the specified type live at the specified time.
   * This function continues to accumulate results until a query returns no results.
   *
   * @param time the "live" time for the retrieval.
   * @param query an interface which performs the type appropriate query call.
   * @param <A> class of the object which will be returned.
   * @return the list of results.
   * @throws IOException on any DB error.
   */
  @SuppressWarnings("Duplicates")
  public static <A extends CachedData> List<A> retrieveAll(long time, QueryCaller<A> query) throws IOException {
    final AttributeSelector ats = makeAtSelector(time);
    long contid = 0;
    List<A> results = new ArrayList<>();
    List<A> nextBatch = query.query(contid, ats);
    while (!nextBatch.isEmpty()) {
      results.addAll(nextBatch);
      contid = nextBatch.get(nextBatch.size() - 1).getCid();
      nextBatch = query.query(contid, ats);
    }
    return results;
  }

}
