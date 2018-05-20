package enterprises.orbital.evekit.snapshot;

import enterprises.orbital.base.OrbitalProperties;
import enterprises.orbital.evekit.account.SynchronizedEveAccount;
import enterprises.orbital.evekit.snapshot.capsuleer.*;
import enterprises.orbital.evekit.snapshot.common.*;
import enterprises.orbital.evekit.snapshot.corporation.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.zip.ZipOutputStream;

/**
 * Create a snapshot of the given account. A separate cron job queues these up every few hours and deletes older snapshots as necessary.
 */
public class SnapshotScheduler {
  protected static final Logger log = Logger.getLogger(SnapshotScheduler.class.getName());
  private static final String PROP_SNAPSHOT_DIR = "enterprises.orbital.evekit.snapshot.directory";
  private static final String DEF_SNAPSHOT_DIR = ".";

  private static String makeSnapshotFileNamePrefix(
      SynchronizedEveAccount acct) {
    return "snapshot_" + acct.getUserAccount()
                             .getUid() + "_" + acct.getAid() + "_";
  }

  private static String makeSnapshotFileName(
      SynchronizedEveAccount acct,
      Date when) {
    SimpleDateFormat formatter = new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss");
    return makeSnapshotFileNamePrefix(acct) + formatter.format(when) + ".zip";
  }

  @SuppressWarnings("WeakerAccess")
  public static List<File> findSnapshotFiles(
      SynchronizedEveAccount acct) {
    // Retrieve all snapshots for this account
    final String prefix = makeSnapshotFileNamePrefix(acct);
    String snapshotDir = OrbitalProperties.getGlobalProperty(PROP_SNAPSHOT_DIR, DEF_SNAPSHOT_DIR);
    File dir = new File(snapshotDir);
    File[] fileList = dir.listFiles((dir1, name) -> name.startsWith(prefix));
    if (fileList == null) return Collections.emptyList();
    // This next copy is necessary because Arrays.asList returns an inner class which does
    // not support the remove method.
    List<File> eligible = new ArrayList<>(Arrays.asList(fileList));
    // Sort ascending by date.
    eligible.sort(new Comparator<File>() {
      final String formatString = "yyyy_MM_dd_HH_mm_ss";
      final SimpleDateFormat formatter = new SimpleDateFormat(formatString);

      @Override
      public int compare(
          File arg0,
          File arg1) {
        try {
          Date f1 = formatter.parse(arg0.getName()
                                        .substring(prefix.length(), prefix.length() + formatString.length()));
          Date f2 = formatter.parse(arg1.getName()
                                        .substring(prefix.length(), prefix.length() + formatString.length()));
          return f1.compareTo(f2);
        } catch (ParseException e) {
          log.log(Level.WARNING, "Failed to compare snapshot file names: " + arg0 + " " + arg1, e);
        }
        return 0;
      }
    });
    return eligible;
  }

  @SuppressWarnings("unused")
  public static long lastSnapshotTime(
      SynchronizedEveAccount acct)
      throws ParseException {
    List<File> snapshots = findSnapshotFiles(acct);
    if (snapshots.size() > 0) {
      String prefix = makeSnapshotFileNamePrefix(acct);
      String formatString = "yyyy_MM_dd_HH_mm_ss";
      SimpleDateFormat formatter = new SimpleDateFormat(formatString);
      File last = snapshots.get(snapshots.size() - 1);
      Date when = formatter.parse(last.getName()
                                      .substring(prefix.length(), prefix.length() + formatString.length()));
      return when.getTime();
    }
    return 0;
  }

  @SuppressWarnings("unused")
  public static void generateAccountSnapshot(
      SynchronizedEveAccount acct,
      long at)
      throws IOException {
    // Generate output zip with one entry per sheet.
    log.info("Generating snapshot for: " + acct);
    Date now = new Date(at);
    String filename = makeSnapshotFileName(acct, now);
    String snapshotDir = OrbitalProperties.getGlobalProperty(PROP_SNAPSHOT_DIR, DEF_SNAPSHOT_DIR);
    ZipOutputStream writer = new ZipOutputStream(new FileOutputStream(snapshotDir + File.separator + filename));
    createAccountDump(acct, writer, at);
    writer.close();
    log.info("Snapshot complete: " + acct);

    // Since we successfully wrote this file. Check for older files and remove.
    List<File> oldFiles = findSnapshotFiles(acct);
    // We want to keep the latest file, so remove that from the list
    if (oldFiles.size() > 0) {
      oldFiles.remove(oldFiles.size() - 1);
    }
    // Delete all older files
    for (File next : oldFiles) {
      if (!next.delete()) log.warning("Failed to delete eligible snapshot file: " + next);
    }
    log.info("Cleaned up " + oldFiles.size() + " files");
  }

  private static void createAccountDump(
      SynchronizedEveAccount toDump,
      ZipOutputStream writer,
      long at)
      throws IOException {

    // Write out common components
    AccountBalanceSheetWriter.dumpToSheet(toDump, writer, at);
    AssetSheetWriter.dumpToSheet(toDump, writer, at);
    BlueprintSheetWriter.dumpToSheet(toDump, writer, at);
    BookmarkSheetWriter.dumpToSheet(toDump, writer, at);
    ContactSheetWriter.dumpToSheet(toDump, writer, at);
    ContactLabelSheetWriter.dumpToSheet(toDump, writer, at);
    ContractSheetWriter.dumpToSheet(toDump, writer, at);
    FacWarStatsSheetWriter.dumpToSheet(toDump, writer, at);
    IndustryJobSheetWriter.dumpToSheet(toDump, writer, at);
    KillSheetWriter.dumpToSheet(toDump, writer, at);
    LocationSheetWriter.dumpToSheet(toDump, writer, at);
    MarketOrderSheetWriter.dumpToSheet(toDump, writer, at);
    StandingSheetWriter.dumpToSheet(toDump, writer, at);
    WalletJournalSheetWriter.dumpToSheet(toDump, writer, at);
    WalletTransactionSheetWriter.dumpToSheet(toDump, writer, at);

    if (toDump.isCharacterType()) {
      // Capsuleers:
      CalendarSheetWriter.dumpToSheet(toDump, writer, at);
      CharacterSheetSheetWriter.dumpToSheet(toDump, writer, at);
      CharacterSheetAttributesSheetWriter.dumpToSheet(toDump, writer, at);
      CharacterSheetCloneSheetWriter.dumpToSheet(toDump, writer, at);
      CharacterSheetJumpSheetWriter.dumpToSheet(toDump, writer, at);
      CharacterSheetSkillPointsSheetWriter.dumpToSheet(toDump, writer, at);
      FittingsSheetWriter.dumpToSheet(toDump, writer, at);
      FleetsSheetWriter.dumpToSheet(toDump, writer, at);
      ImplantSheetWriter.dumpToSheet(toDump, writer, at);
      JumpCloneSheetWriter.dumpToSheet(toDump, writer, at);
      JumpCloneImplantSheetWriter.dumpToSheet(toDump, writer, at);
      LoyaltyPointsSheetWriter.dumpToSheet(toDump, writer, at);
      ChatChannelSheetWriter.dumpToSheet(toDump, writer, at);
      ContactNotificationSheetWriter.dumpToSheet(toDump, writer, at);
      MailingListSheetWriter.dumpToSheet(toDump, writer, at);
      MailLabelSheetWriter.dumpToSheet(toDump, writer, at);
      MailMessageSheetWriter.dumpToSheet(toDump, writer, at);
      MiningLedgerSheetWriter.dumpToSheet(toDump, writer, at);
      MedalSheetWriter.dumpToSheet(toDump, writer, at);
      NotificationSheetWriter.dumpToSheet(toDump, writer, at);
      OpportunitySheetWriter.dumpToSheet(toDump, writer, at);
      PlanetaryColoniesSheetWriter.dumpToSheet(toDump, writer, at);
      ResearchAgentSheetWriter.dumpToSheet(toDump, writer, at);
      CharacterRoleSheetWriter.dumpToSheet(toDump, writer, at);
      SkillSheetWriter.dumpToSheet(toDump, writer, at);
      SkillsInQueueSheetWriter.dumpToSheet(toDump, writer, at);
      TitleSheetWriter.dumpToSheet(toDump, writer, at);
      CharacterLocationSheetWriter.dumpToSheet(toDump, writer, at);
      CharacterShipSheetWriter.dumpToSheet(toDump, writer, at);
      CharacterOnlineSheetWriter.dumpToSheet(toDump, writer, at);
    } else {
      // Corporations:
      ContainerLogSheetWriter.dumpToSheet(toDump, writer, at);
      CorporationDivisionSheetWriter.dumpToSheet(toDump, writer, at);
      CorporationMedalSheetWriter.dumpToSheet(toDump, writer, at);
      CorporationMemberMedalSheetWriter.dumpToSheet(toDump, writer, at);
      CorporationSheetSheetWriter.dumpToSheet(toDump, writer, at);
      CorporationTitleSheetWriter.dumpToSheet(toDump, writer, at);
      CorporationTitleRoleSheetWriter.dumpToSheet(toDump, writer, at);
      CustomsOfficesSheetWriter.dumpToSheet(toDump, writer, at);
      FacilitiesSheetWriter.dumpToSheet(toDump, writer, at);
      MemberSheetWriter.dumpToSheet(toDump, writer, at);
      MemberRoleSheetWriter.dumpToSheet(toDump, writer, at);
      MemberRoleHistorySheetWriter.dumpToSheet(toDump, writer, at);
      MemberTitleSheetWriter.dumpToSheet(toDump, writer, at);
      MemberTrackingSheetWriter.dumpToSheet(toDump, writer, at);
      MiningExtractionSheetWriter.dumpToSheet(toDump, writer, at);
      MiningObserverSheetWriter.dumpToSheet(toDump, writer, at);
      MiningObservationSheetWriter.dumpToSheet(toDump, writer, at);
      ShareholderSheetWriter.dumpToSheet(toDump, writer, at);
      StarbaseSheetWriter.dumpToSheet(toDump, writer, at);
      StructureSheetWriter.dumpToSheet(toDump, writer, at);
    }

  }

}
