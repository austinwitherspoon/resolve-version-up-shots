"""Resolve Version Up Shots
A tool to help manage versions (usually in a VFX context).
"""

import contextlib
import re
import glob
import logging
from pathlib import Path
import sys
import DaVinciResolveScript  # type: *


logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter(
    "[%(asctime)s] %(levelname)s - %(message)s",
    "%m-%d %H:%M:%S",
)
handler = logging.StreamHandler()
handler.setFormatter(formatter)
logger.addHandler(handler)

# Regex to find all version number strings (like v0001)
# but only match to the last one in each part of the directory.
# If there are multiple 'versions' in the path, it will ignore everything
# but the last one in the file name and each directory folder.
VERSION_SCAN_REGEX = re.compile(
    r"([vV]{1}[0-9]+)(?![^\\/]+[\\/]+)(?![^\\/]+[vV]{1}\d+)"
)

# Regex to pull out the frame range from resolve's file names
FRAME_SEQUENCE_REGEX = re.compile(r"(\[[0-9-]+\])")


RESOLVE = DaVinciResolveScript.scriptapp("Resolve")
PROJECT = RESOLVE.GetProjectManager().GetCurrentProject()
FUSION = RESOLVE.Fusion()
UI = FUSION.UIManager
DISPATCHER = DaVinciResolveScript.UIDispatcher(UI)


class VersionUpShotsWindow:
    winID = "com.austinwitherspoon.resolve.VersionUpShots"
    _scanning = False

    def __init__(self):
        self.shots = []
        self.project = PROJECT

        self.create_ui()
        self.build_shot_list()
        self.populate_track_list()

        self.window.Show()

    def create_ui(self):
        """Create the primary UI for the tool.
        This is using Resolve's weird custom QT(?) version."""

        # If the window is already open, cancel this and bring it to the front.
        # Only one instance should exist at a time.
        self.window = UI.FindWindow(self.winID)
        if self.window:
            self._scanning = False
            self.window.Show()
            self.window.Raise()
            exit()

        # Build interface
        self.window = DISPATCHER.AddWindow(
            {
                "ID": self.winID,
                "Geometry": [100, 100, 500, 700],
                "WindowTitle": "Version Up Shots",
            },
            UI.VGroup(
                [
                    UI.Label(
                        {
                            "Text": "VFX Version Up Shots Tool",
                            "Weight": 0,
                            "Font": UI.Font({"Family": "Verdana", "PixelSize": 20}),
                        }
                    ),
                    UI.VGap(5, 0),
                    UI.Label({"ID": "Status", "Weight": 0}),
                    UI.ComboBox({"ID": "Track", "Text": "Video Track"}),
                    UI.Button(
                        {
                            "ID": "ScanVersions",
                            "Text": "Scan For Latest Versions",
                            "Weight": 0,
                        }
                    ),
                    UI.VGap(15, 0),
                    UI.Label({"Text": "Shots", "Weight": 0}),
                    UI.Tree({"ID": "ShotTree"}),
                    UI.VGap(15, 0),
                    UI.Label({"Text": "Where to import new shots:", "Weight": 0}),
                    UI.ComboBox({"ID": "Location"}),
                    UI.VGap(15, 0),
                    UI.Button(
                        {"ID": "Submit", "Text": "Update Versions!", "Weight": 0}
                    ),
                    UI.Label(
                        {
                            "Text": "Made by Austin Witherspoon ♥",
                            "Weight": 0,
                            "Font": UI.Font({"PixelSize": 10}),
                        }
                    ),
                ]
            ),
        )

        self.window.Find("Location").AddItems(
            ["Currently Open Bin", "Same Bin As Original Clip"]
        )

        # Register Events
        self.window.On[self.winID].Close = self.closeEvent
        self.window.On["ScanVersions"].Clicked = self.scan_versions
        self.window.On["Submit"].Clicked = self.version_up_shots

    def version_up_shots(self, event):
        """Swap out all versions with the highest available version."""
        if self._scanning:
            return

        if not self.shots:
            return

        tree = self.window.Find("ShotTree")

        importToSourceBin = (
            self.window.Find("Location").CurrentText == "Same Bin As Original Clip"
        )

        i = 0
        for shot in self.shots:
            success = shot.update(importToSourceBin)
            row = tree.TopLevelItem(i)
            if success:
                row.Text[1] = shot.highestVersion
            else:
                row.Text[0] = "!! FAILED !! " + row.Text[0]
            i += 1

    def populate_track_list(self):
        dropdown = self.window.Find("Track")
        timeline = self.project.GetCurrentTimeline()

        videoTracks = [
            timeline.GetTrackName("video", i + 1)
            for i in range(timeline.GetTrackCount("video"))
        ]
        videoTracks.append("All Tracks")
        videoTracks.reverse()

        dropdown.AddItems(videoTracks)

    def scan_versions(self, event):
        if self._scanning:
            return

        logger.info("Scanning Versions..")

        self._scanning = True

        self.window.Find("Status").SetText("Scanning versions..")

        timeline = self.project.GetCurrentTimeline()
        logger.debug(f"Timeline: {timeline.GetName()}")

        targetTrack = self.window.Find("Track").CurrentText
        logger.debug(f"Target Track: {targetTrack}")

        # Grab all clips in selected track
        clips = []
        if targetTrack == "All Tracks":
            logger.info("Scanning all tracks..")
            for track in [
                (i + 1, timeline.GetTrackName("video", i + 1))
                for i in range(timeline.GetTrackCount("video"))
            ]:
                index, name = track
                logger.info(f"Scanning track {index} ({name})")
                found = timeline.GetItemListInTrack("video", index)
                logger.info(f"Found {len(found)} clips")
                clips += found

        else:
            for track in [
                (i + 1, timeline.GetTrackName("video", i + 1))
                for i in range(timeline.GetTrackCount("video"))
            ]:
                index, name = track
                if name == targetTrack:
                    logger.info(f"Scanning track {index} ({name})")
                    clips += timeline.GetItemListInTrack("video", index)
                    logger.info(f"Found {len(clips)} clips")
                    break

        self.shots = []

        for clip in clips:
            clip_name = clip.GetName()
            logger.info("Scanning " + clip_name)
            self.window.Find("Status").SetText("Scanning " + clip_name)

            with contextlib.suppress(AttributeError):
                shot = Shot(clip)
                self.shots.append(shot)

        logger.debug("Building shot list..")
        self.build_shot_list()

        bad = [i for i in self.shots if i.highestVersion != i.highestInvalidVersion]

        if len(bad) > 0:
            self.alert(bad)

        logger.info("Done!")
        self.window.Find("Status").SetText("")
        self._scanning = False

    def alert(self, shots):
        alert = DISPATCHER.AddWindow(
            {
                "ID": self.winID + "Error",
                "Geometry": [100, 100, 500, 300],
                "WindowTitle": "Warning!",
            },
            UI.VGroup(
                [
                    UI.Label(
                        {
                            "Text": "Warning!",
                            "Weight": 0,
                            "Font": UI.Font({"Family": "Verdana", "PixelSize": 20}),
                        }
                    ),
                    UI.VGap(20, 0),
                    UI.Label(
                        {
                            "Text": "Some shots had latest versions that were either missing frames or not long enough. \nUsing the highest available version that fits the frame range on these clips for now.",
                            "WordWrap": True,
                            "Weight": 0,
                        }
                    ),
                ]
            ),
        )
        alert.On[self.winID + "Error"].Close = lambda ev: alert.Hide()

        alert.Show()

    def build_shot_list(self):
        tree = self.window.Find("ShotTree")
        tree.Clear()

        header = tree.NewItem()
        header.Text[0] = "Shot"
        header.Text[1] = "Current"
        header.Text[2] = "Highest"
        header.Text[3] = "Highest Valid Render"
        tree.SetHeaderItem(header)

        tree.ColumnCount = 4
        tree.ColumnWidth[0] = 210
        tree.ColumnWidth[1] = 70
        tree.ColumnWidth[2] = 70
        tree.ColumnWidth[3] = 100

        if self.shots:
            for shot in self.shots:
                row = tree.NewItem()
                row.Text[0] = shot.name
                row.Text[1] = Path(shot.currentVersion).name
                row.Text[2] = (
                    Path(shot.highestInvalidVersion).name
                    if shot.highestInvalidVersion
                    else ""
                )
                row.Text[3] = Path(shot.highestVersion).name
                tree.AddTopLevelItem(row)

    def closeEvent(self, event):
        DISPATCHER.ExitLoop()


class Shot:
    name = None
    trackItem = None
    currentVersion = None
    highestVersion = None
    highestInvalidVersion = None
    isVersionable = False
    mpItem = None
    path = None
    duration = None
    isSequence = False
    invalidVersions = None

    def __init__(self, trackItem):
        global VERSION_SCAN_REGEX
        self.name = trackItem.GetName()
        logger.debug(f"Creating new shot object for {self.name}")
        self.trackItem = trackItem
        self.duration = trackItem.GetDuration()
        self.invalidVersions = []

        self.mpItem = trackItem.GetMediaPoolItem()
        self.path = self.mpItem.GetClipProperty("File Path")

        logger.debug("Path: " + self.path)

        # restrict to unique entries. Proper version paths should be left with 1 unique version
        version = list(set(VERSION_SCAN_REGEX.findall(self.path)))

        logger.debug(
            f"Extracted version from clip: {VERSION_SCAN_REGEX.findall(self.path)}"
        )

        if len(version) == 0:
            self.isVersionable = False
            self.currentVersion = self.path
            self.highestVersion = self.path
            return
        self.isVersionable = True
        self.currentVersion = version[-1]

        available = self.availableVersions()

        self.highestVersion = available[-1]

    def availableVersions(self):
        global versionRegex, frameSequenceRegex

        version = list(set(VERSION_SCAN_REGEX.findall(self.path)))[-1]
        globPath = self.path.replace(version, "*")

        frameRange = FRAME_SEQUENCE_REGEX.findall(globPath)

        logger.debug(
            f"Scanning {self.name}.. version: {version}.\n\tGlob: {globPath}\n\tFrame Range: {frameRange}"
        )

        self.isSequence = isSequence = len(frameRange) > 0
        if isSequence:
            # If we have a sequence in a similarly named folder
            if len(globPath.split("*")) > 2:
                # get rid of the file + extension so we run faster
                globPath = re.sub(r"([^\\\/]+)$", "", globPath)
            else:
                globPath = re.sub(FRAME_SEQUENCE_REGEX, "*", globPath)

        logger.debug("Glob: " + globPath)
        logger.debug("is sequence?" + str(isSequence))

        results = sorted(
            list(
                set(
                    [
                        VERSION_SCAN_REGEX.findall(i)[-1]
                        for i in sorted(glob.glob(globPath))
                    ]
                )
            )
        )

        logger.debug("results:")
        logger.debug(results)

        validResults = self.validateVersions(results)

        self.invalidVersions = [i for i in results if i not in validResults]

        self.highestInvalidVersion = results[-1]

        return validResults

    def validateVersions(self, versions):
        if not self.isSequence:
            return versions

        toScan = versions[:]
        toScan.reverse()

        goodVersions = []

        foundGoodRender = False

        for version in toScan:
            versionPath = self.path.replace(self.currentVersion, version)
            sequence = FRAME_SEQUENCE_REGEX.findall(versionPath)[-1]
            globPath = versionPath.replace(sequence, "*")

            files = glob.glob(globPath)
            frames = sorted(
                [
                    i.replace(globPath.split("*")[0], "").replace(
                        globPath.split("*")[1], ""
                    )
                    for i in files
                ]
            )

            if not self.missingFrames(frames) and len(frames) >= self.duration:
                goodVersions.append(version)
                foundGoodRender = True
                break

        return goodVersions

    def missingFrames(self, frames):
        if len(frames) == 0:
            return True
        start = int(frames[0])
        i = start
        for frame in frames:
            if i == int(frame):
                i += 1
            else:
                return True
        return False

    def update(self, importToSourceBin=False):
        global RESOLVE, PROJECT
        ms = RESOLVE.GetMediaStorage()

        if self.currentVersion == self.highestVersion:
            return True

        newPath = self.path.replace(self.currentVersion, self.highestVersion)
        if self.isSequence:
            newPath = newPath.replace(
                re.findall(r"[\\\/]{1}([^\\\/]+)(?!.+)", newPath)[-1], ""
            )

        if importToSourceBin:
            folder = self.findFolder(self.mpItem)
            PROJECT.GetMediaPool().SetCurrentFolder(folder)

        ms.AddItemListToMediaPool(newPath)

        item = self.findItemInProject(newPath)

        if not item:
            logger.error(f"Could not find {newPath} in project!")
            return False

        self.swap(item)

        if newPath in self.trackItem.GetMediaPoolItem().GetClipProperty("File Path"):
            logger.info(f"Successfully updated {self.name} to {newPath}")
            return True

        logger.error(f"Failed to update {self.name} to {newPath}")

        return False

    def swap(self, mediaPoolItem):
        timelineItem = self.trackItem

        leftOffset = timelineItem.GetLeftOffset()
        start = int(timelineItem.GetMediaPoolItem().GetClipProperty("Start"))
        newIn = start + leftOffset

        timelineItem.AddTake(mediaPoolItem, newIn, newIn + self.duration)
        timelineItem.SelectTakeByIndex(timelineItem.GetTakesCount())
        timelineItem.FinalizeTake()

    def findItemInProject(self, path, folder=None):
        global PROJECT
        if not folder:
            folder = PROJECT.GetMediaPool().GetRootFolder()

        matches = [
            i for i in folder.GetClipList() if path in i.GetClipProperty("File Path")
        ]
        if len(matches) > 0:
            return matches[-1]

        for subfolder in folder.GetSubFolderList():
            result = self.findItemInProject(path, subfolder)
            if result:
                return result

        return False

    def findFolder(self, item, folder=None):
        global PROJECT
        if not folder:
            folder = PROJECT.GetMediaPool().GetRootFolder()

        matches = [
            i
            for i in folder.GetClipList()
            if i.GetClipProperty() == item.GetClipProperty()
        ]
        if len(matches) > 0:
            return folder

        for subfolder in folder.GetSubFolderList():
            result = self.findFolder(item, subfolder)
            if result:
                return result

        return False


# No `if __name__ == "__main__"` because this is a Resolve script,
# resolve doesn't run as `__main__``  when launched from the Scripts menu
VersionUpShotsWindow()
DISPATCHER.RunLoop()
