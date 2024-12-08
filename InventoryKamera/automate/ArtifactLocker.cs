using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using static InventoryKamera.Artifact;
using Accord.Imaging;
using System.Text.RegularExpressions;


namespace InventoryKamera.automate
{
    internal class ArtifactLocker : ArtifactScraper
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        private const string artifactOverridesFilePath = @"overrides\artifacts-dec07.json";

        private List<Artifact> artifactOverrides;

        public ArtifactLocker()
        {
            if (File.Exists(artifactOverridesFilePath))
            {
                var json = File.ReadAllText(artifactOverridesFilePath);
                artifactOverrides = JsonConvert.DeserializeObject<List<Artifact>>(json);
            }
            else
            {
                Logger.Error("Artifact file not found.");
                artifactOverrides = new List<Artifact>();
            }
        }

        public void SyncArtifactLocks(int count = 0)
        {
            int artifactCount = count == 0 ? ScanItemCount() : count;
            int page = 1;
            var (rectangles, cols, rows) = GetPageOfItems(page);
            int fullPage = cols * rows;
            int totalRows = (int)Math.Ceiling(artifactCount / (decimal)cols);
            int cardsQueued = 0;
            int rowsQueued = 0;
            UserInterface.SetArtifact_Max(artifactCount);

            StopScanning = false;

            Logger.Info("Found {0} for artifact count.", artifactCount);

            while (cardsQueued < artifactCount)
            {
                Logger.Debug("Scanning artifact page {0}", page);
                Logger.Debug("Located {0} possible item locations on page.", rectangles.Count);

                int cardsRemaining = artifactCount - cardsQueued;
                for (int i = cardsRemaining < fullPage ? (rows - (totalRows - rowsQueued)) * cols : 0; i < rectangles.Count; i++)
                {
                    Rectangle item = rectangles[i];
                    Navigation.SetCursor(item.Center().X, item.Center().Y);
                    Navigation.Click();
                    Navigation.SystemWait(Navigation.Speed.SelectNextInventoryItem);

                    if (ProcessArtifact(cardsQueued))
                    {
                        Navigation.SetCursor(item.Center().X, item.Center().Y);
                        Navigation.Click();
                        Navigation.SystemWait(Navigation.Speed.SelectNextInventoryItem);
                    }
                    cardsQueued++;
                    if (cardsQueued >= artifactCount || StopScanning)
                    {
                        if (StopScanning) Logger.Info("Stopping artifact sync based on filtering");
                        else Logger.Info("Stopping artifact sync based on syncs queued ({0} of {1})", cardsQueued, artifactCount);
                        return;
                    }
                }

                Logger.Debug("Finished syncing page of artifacts. Scrolling...");

                rowsQueued += rows;

                if (totalRows - rowsQueued <= rows)
                {
                    for (int i = 0; i < 10 * (totalRows - rowsQueued) - 1; i++)
                    {
                        Navigation.sim.Mouse.VerticalScroll(-1);
                        Navigation.Wait(1);
                    }
                    Navigation.SystemWait(Navigation.Speed.Fast);
                }
                else
                {
                    for (int i = 0; i < 10 * rows - 1; i++)
                    {
                        Navigation.sim.Mouse.VerticalScroll(-1);
                        Navigation.Wait(1);
                    }
                    if (page % 12 == 0)
                    {
                        Logger.Debug("Scrolled back one");
                        Navigation.sim.Mouse.VerticalScroll(1);
                        Navigation.Wait(1);
                    }
                    Navigation.SystemWait(Navigation.Speed.Fast);
                }
                ++page;
                (rectangles, cols, rows) = GetPageOfItems(page, acceptLess: totalRows - rowsQueued <= fullPage);
            }
        }

        private void ClearFilters()
        {

            using (var x = Navigation.CaptureRegion(
                x: (int)((Navigation.IsNormal ? 0.0750 : 0.0757) * Navigation.GetWidth()),
                y: (int)((Navigation.IsNormal ? 0.8522 : 0.8678) * Navigation.GetHeight()),
                width: (int)((Navigation.IsNormal ? 0.2244 : 0.2236) * Navigation.GetWidth()),
                height: (int)((Navigation.IsNormal ? 0.0422 : 0.0367) * Navigation.GetHeight())))
            {
                //Navigation.DisplayBitmap(x);
                var t = GenshinProcesor.AnalyzeText(x).Trim().ToLower();
                if (t != null && t.Contains("filter"))
                {
                    Navigation.ClearArtifactFilters();
                }
            }
        }

        private bool ProcessArtifact(int id)
        {
            var lockStatusChanged = false;
            var card = GetItemCard();

            Bitmap name = GetItemNameBitmap(card);
            Bitmap locked = GetLockedBitmap(card);
            Bitmap equipped = GetEquippedBitmap(card);
            Bitmap gearSlot = GetGearSlotBitmap(card);
            Bitmap mainStat = GetMainStatBitmap(card);
            Bitmap level = GetLevelBitmap(card);
            Bitmap subStats = GetSubstatsBitmap(card);

            List<Bitmap> bitmaps = new List<Bitmap> { name, gearSlot, mainStat, level, subStats, equipped, locked };

            Artifact artifact = GetArtifactDataFromBitmaps(bitmaps, id);

            bool? overrideStatus = getOverrideStatus(artifact);
            if (overrideStatus != null)
            {
                if (overrideStatus != artifact.Lock)
                {
                    ToggleLockStatus();
                    Logger.Info("Artifact {0} locked status overridden to {1}", artifact.SetName, overrideStatus);
                    lockStatusChanged = true;
                }
            }

            bitmaps.ForEach(b => b.Dispose());
            card.Dispose();
            return lockStatusChanged;
        }

        private bool? getOverrideStatus(Artifact artifactOnCard)
        {
            foreach (var artifactOverride in artifactOverrides)
            {
                if (artifactOnCard.IsSameAs(artifactOverride))
                {
                    return artifactOverride.Lock;
                }
            }
            return null; // Return null if no matching artifact override is found
        }

        private void ToggleLockStatus()
        {
            // 0.8643
            Navigation.SetCursor((int)(Navigation.GetWidth() * 0.883), (int)(Navigation.GetHeight() * 0.4085));
            Navigation.Click();
            Navigation.SystemWait(Navigation.Speed.SelectNextInventoryItem);
        }


        private Bitmap GetSubstatsBitmap(Bitmap card)
        {
            return GenshinProcesor.CopyBitmap(card, new Rectangle(
                x: (int)(card.Width * 0.0911),
                y: (int)(card.Height * (Navigation.IsNormal ? 0.4216 : 0.3682)),
                width: (int)(card.Width * 0.8097),
                height: (int)(card.Height * (Navigation.IsNormal ? 0.1841 : 0.1573))));
        }

        private Bitmap GetMainStatBitmap(Bitmap card)
        {
            return GenshinProcesor.CopyBitmap(card, new Rectangle(
                x: (int)(card.Width * 0.0405),
                y: (int)(card.Height * (Navigation.IsNormal ? 0.1722 : 0.1477)),
                width: (int)(card.Width * 0.4555),
                height: (int)(card.Height * (Navigation.IsNormal ? 0.0416 : 0.0416))));
        }

        private Bitmap GetLevelBitmap(Bitmap card)
        {
            return GenshinProcesor.CopyBitmap(card, new Rectangle(
                x: (int)(card.Width * 0.0506),
                y: (int)(card.Height * (Navigation.IsNormal ? 0.3634 : 0.3197)),
                width: (int)(card.Width * 0.1417),
                height: (int)(card.Height * (Navigation.IsNormal ? 0.0416 : 0.0347))));
        }

        private Bitmap GetGearSlotBitmap(Bitmap card)
        {
            return GenshinProcesor.CopyBitmap(card, new Rectangle(
                x: (int)(card.Width * 0.0405),
                y: (int)(card.Height * (Navigation.IsNormal ? 0.07720 : 0.0663)),
                width: (int)(card.Width * 0.4757),
                height: (int)(card.Height * (Navigation.IsNormal ? 0.0475 : 0.0809))));
        }

        private static Artifact GetArtifactDataFromBitmaps(List<Bitmap> bm, int id)
        {
            // Init Variables
            string gearSlot = null;
            string mainStat = null;
            string setName = null;
            string equippedCharacter = null;
            List<SubStat> subStats = new List<SubStat>();
            int rarity = 0;
            int level = 0;
            bool _lock = false;

            if (bm.Count >= 6)
            {
                int a_name = 0; int a_gearSlot = 1; int a_mainStat = 2; int a_level = 3; int a_subStats = 4; int a_lock = 6;
                // Get Rarity
                rarity = GetRarity(bm[a_name]);

                // Check for lock color
                Color lockedColor = Color.FromArgb(255, 70, 80, 100); // Dark area around red lock
                Color lockStatus = bm[a_lock].GetPixel(10, 10);
                _lock = GenshinProcesor.CompareColors(lockedColor, lockStatus);

                // Improved Scanning using multi threading
                List<Task> tasks = new List<Task>();

                var taskGear = Task.Run(() => gearSlot = ScanArtifactGearSlot(bm[a_gearSlot]));
                var taskMain = taskGear.ContinueWith((antecedent) => mainStat = ScanArtifactMainStat(bm[a_mainStat], antecedent.Result));
                var taskLevel = Task.Run(() => level = ScanArtifactLevel(bm[a_level]));
                var taskSubs = Task.Run(() => subStats = ScanArtifactSubStats(bm[a_subStats]));
                var taskName = Task.Run(() => setName = ScanArtifactSet(bm[a_name]));

                tasks.Add(taskGear);
                tasks.Add(taskMain);
                tasks.Add(taskLevel);
                tasks.Add(taskSubs);
                tasks.Add(taskName);

                Task.WaitAll(tasks.ToArray());
            }
            return new Artifact(setName, rarity, level, gearSlot, mainStat, subStats, equippedCharacter, id, _lock);
        }

        private static int GetRarity(Bitmap bm)
        {
            var averageColor = new ImageStatistics(bm);

            Color fiveStar = Color.FromArgb(255, 188, 105, 50);
            Color fourStar = Color.FromArgb(255, 161, 86, 224);
            Color threeStar = Color.FromArgb(255, 81, 127, 203);
            Color twoStar = Color.FromArgb(255, 42, 143, 114);
            Color oneStar = Color.FromArgb(255, 114, 119, 138);

            var colors = new List<Color> { Color.Black, oneStar, twoStar, threeStar, fourStar, fiveStar };

            var c = GenshinProcesor.ClosestColor(colors, averageColor);

            return colors.IndexOf(c);
        }

        private static string ScanEnhancementMaterialName(Bitmap bm)
        {
            GenshinProcesor.SetGamma(0.2, 0.2, 0.2, ref bm);
            Bitmap n = GenshinProcesor.ConvertToGrayscale(bm);
            GenshinProcesor.SetInvert(ref n);

            // Analyze
            string name = Regex.Replace(GenshinProcesor.AnalyzeText(n).ToLower(), @"[\W]", string.Empty);
            name = GenshinProcesor.FindClosestMaterialName(name);
            n.Dispose();

            return name;
        }

        #region Task Methods

        private static string ScanArtifactGearSlot(Bitmap bm)
        {
            // Process Img
            Bitmap n = GenshinProcesor.ConvertToGrayscale(bm);
            GenshinProcesor.SetContrast(80.0, ref n);
            GenshinProcesor.SetInvert(ref n);

            string gearSlot = GenshinProcesor.AnalyzeText(n).Trim().ToLower();
            gearSlot = Regex.Replace(gearSlot, @"[\W_]", string.Empty);
            gearSlot = GenshinProcesor.FindClosestGearSlot(gearSlot);
            n.Dispose();
            return gearSlot;
        }

        private static string ScanArtifactMainStat(Bitmap bm, string gearSlot)
        {
            switch (gearSlot)
            {
                // Flower of Life. Flat HP
                case "flower":
                    return GenshinProcesor.Stats["hp"];

                // Plume of Death. Flat ATK
                case "plume":
                    return GenshinProcesor.Stats["atk"];

                // Otherwise it's either sands, goblet or circlet.
                default:
                    Bitmap copy = (Bitmap)bm.Clone();
                    GenshinProcesor.SetContrast(100.0, ref copy);
                    Bitmap n = GenshinProcesor.ConvertToGrayscale(copy);

                    GenshinProcesor.SetThreshold(135, ref n);
                    GenshinProcesor.SetInvert(ref n);

                    // Get Main Stat
                    string mainStat = GenshinProcesor.AnalyzeText(n).ToLower().Trim();


                    // Remove anything not a-z as well as removes spaces/underscores
                    mainStat = Regex.Replace(mainStat, @"[\W_0-9]", string.Empty);

                    mainStat = GenshinProcesor.FindClosestStat(mainStat, 80);

                    if (mainStat == "def" || mainStat == "atk" || mainStat == "hp")
                    {
                        mainStat += "_";
                    }
                    n.Dispose();
                    copy.Dispose();
                    return mainStat;
            }
        }

        private static int ScanArtifactLevel(Bitmap bm)
        {
            // Process Img
            Bitmap n = GenshinProcesor.ConvertToGrayscale(bm);
            GenshinProcesor.SetContrast(80.0, ref n);
            GenshinProcesor.SetInvert(ref n);

            // numbersOnly = true => seems to interpret the '+' as a '4'
            string text = GenshinProcesor.AnalyzeText(n, Tesseract.PageSegMode.SingleWord).Trim().ToLower();
            n.Dispose();

            // Get rid of all non digits
            text = Regex.Replace(text, @"[\D]", string.Empty);

            return int.TryParse(text, out int level) ? level : -1;
        }

        private static List<SubStat> ScanArtifactSubStats(Bitmap artifactImage)
        {
            Bitmap bm = (Bitmap)artifactImage.Clone();
            List<string> lines = new List<string>();
            List<SubStat> substats = new List<SubStat>();
            string text;
            GenshinProcesor.SetBrightness(-30, ref bm);
            GenshinProcesor.SetContrast(85, ref bm);
            using (var n = GenshinProcesor.ConvertToGrayscale(bm))
            {
                text = GenshinProcesor.AnalyzeText(n, Tesseract.PageSegMode.Auto).ToLower();
            }

            lines = new List<string>(text.Split('\n'));
            lines.RemoveAll(line => string.IsNullOrWhiteSpace(line));

            var index = lines.FindIndex(line => line.Contains(":") || line.Contains("piece") || line.Contains("set") || line.Contains("2-"));
            if (index >= 0)
            {
                lines.RemoveRange(index, lines.Count - index);
            }

            bm.Dispose();
            for (int i = 0; i < lines.Count; i++)
            {
                var line = Regex.Replace(lines[i], @"(?:^[^a-zA-Z]*)", string.Empty).Replace(" ", string.Empty);

                if (line.Any(char.IsDigit))
                {
                    Logger.Debug("Parsing artifact substat: {0}", line);

                    SubStat substat = new SubStat();
                    Regex re = new Regex(@"^(.*?)(\d+.*)");
                    var result = re.Match(line);
                    var stat = Regex.Replace(result.Groups[1].Value, @"[^\w]", string.Empty);
                    var value = result.Groups[2].Value;

                    string name = line.Contains("%") ? stat + "%" : stat;

                    substat.stat = GenshinProcesor.FindClosestStat(name, 80) ?? "";

                    // Remove any non digits.
                    value = Regex.Replace(value, @"[^0-9]", string.Empty);

                    // Try to parse number
                    if (!decimal.TryParse(value, out substat.value))
                    {
                        Logger.Debug("Failed to parse stat value from: {0}", line);
                        substat.value = -1;
                    }

                    if (substat.value != -1 && substat.stat.Contains("_"))
                    {
                        substat.value /= 10;
                    }

                    if (string.IsNullOrWhiteSpace(substat.stat) || substat.value == -1)
                    {
                        Logger.Debug("Failed to parse stat from: {0}", line);
                    }

                    substats.Insert(i, substat);
                }
            }
            return substats;
        }

        private static string ScanArtifactEquippedCharacter(Bitmap bm)
        {
            Bitmap n = GenshinProcesor.ConvertToGrayscale(bm);
            GenshinProcesor.SetContrast(60.0, ref n);

            string equippedCharacter = GenshinProcesor.AnalyzeText(n).ToLower();
            n.Dispose();

            if (equippedCharacter != "")
            {
                if (equippedCharacter.Contains("equipped") && equippedCharacter.Contains(":"))
                {
                    equippedCharacter = Regex.Replace(equippedCharacter.Split(':')[1], @"[\W]", string.Empty);
                    equippedCharacter = GenshinProcesor.FindClosestCharacterName(equippedCharacter);

                    return equippedCharacter;
                }
            }
            // artifact has no equipped character
            return null;
        }

        private static string ScanArtifactSet(Bitmap itemName)
        {
            GenshinProcesor.SetGamma(0.2, 0.2, 0.2, ref itemName);
            Bitmap grayscale = GenshinProcesor.ConvertToGrayscale(itemName);
            GenshinProcesor.SetInvert(ref grayscale);

            // Analyze
            using (Bitmap padded = new Bitmap((int)(grayscale.Width + grayscale.Width * .1), grayscale.Height + (int)(grayscale.Height * .5)))
            {
                using (Graphics g = Graphics.FromImage(padded))
                {
                    g.Clear(Color.White);
                    g.DrawImage(grayscale, (padded.Width - grayscale.Width) / 2, (padded.Height - grayscale.Height) / 2);

                    var scannedText = GenshinProcesor.AnalyzeText(grayscale, Tesseract.PageSegMode.Auto).ToLower().Replace("\n", " ");
                    string text = Regex.Replace(scannedText, @"[\W]", string.Empty);
                    text = GenshinProcesor.FindClosestArtifactSetFromArtifactName(text);

                    grayscale.Dispose();

                    return text;
                }
            }
        }

        #endregion Task Methods
    }
}
