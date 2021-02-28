#!/usr/bin/env python
#--coding: utf-8 --
"""
Austin's Bulk Timeline Exporter
A tool to help export a large number of timelines quickly.
"""

import sys
import re
import glob

# Regex to find all version number strings (like v0001)
# but only match to the last one in each part of the directory.
# If there are multiple 'versions' in the path, it will ignore everything 
# but the last one in the file name and each directory folder.
versionRegex = re.compile(r'([vV]{1}[0-9]+)(?![^\\/]+[\\/]+)(?![^\\/]+[vV]{1})')

frameSequenceRegex = re.compile(r'(\[[0-9-]+\])')

# Stolen from python_get_resolve.py in the examples folder.
def GetResolve():
    try:
    # The PYTHONPATH needs to be set correctly for this import statement to work.
    # An alternative is to import the DaVinciResolveScript by specifying absolute path (see ExceptionHandler logic)
        import DaVinciResolveScript as bmd
    except ImportError:
        if sys.platform.startswith("darwin"):
            expectedPath="/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules/"
        elif sys.platform.startswith("win") or sys.platform.startswith("cygwin"):
            import os
            expectedPath=os.getenv('PROGRAMDATA') + "\\Blackmagic Design\\DaVinci Resolve\\Support\\Developer\\Scripting\\Modules\\"
        elif sys.platform.startswith("linux"):
            expectedPath="/opt/resolve/libs/Fusion/Modules/"

        # check if the default path has it...
        print("Unable to find module DaVinciResolveScript from $PYTHONPATH - trying default locations")
        try:
            import imp
            bmd = imp.load_source('DaVinciResolveScript', expectedPath+"DaVinciResolveScript.py")
        except ImportError:
            # No fallbacks ... report error:
            print("Unable to find module DaVinciResolveScript - please ensure that the module DaVinciResolveScript is discoverable by python")
            print("For a default DaVinci Resolve installation, the module is expected to be located in: "+expectedPath)
            sys.exit()

    return bmd.scriptapp("Resolve")




resolve = GetResolve()
project = resolve.GetProjectManager().GetCurrentProject()
fusion = resolve.Fusion()
ui = fusion.UIManager



class VersionUpShots:
	winID = 'com.austinwitherspoon.resolve.VersionUpShots'
	window = None
	dispatcher = None
	project = None
	shots = None


	def __init__(self):
		global resolve
		self.shots = []
		self.buildUI()
		self.buildShotList()

		self.project = resolve.GetProjectManager().GetCurrentProject()

		self.populateTrackList()

		# Show window
		self.window.Show()
		self.dispatcher.RunLoop()


	# Build the interface
	def buildUI(self):
		global resolve, fusion, ui, bmd

		# If the window is already open, cancel this and bring it to the front.
		# Only one instance should exist at a time.
		self.window = ui.FindWindow(self.winID)
		if self.window:
			self.window.Show()
			self.window.Raise()
			exit()

		self.dispatcher = bmd.UIDispatcher(ui)

		# Build interface
		self.window = self.dispatcher.AddWindow({
			'ID': self.winID,
			'Geometry': [ 100,100,500,700 ],
			'WindowTitle': "Version Up Shots",
			},
			ui.VGroup([
				ui.Label({ 'Text': "VFX Version Up Shots Tool", 'Weight':0, 'Font': ui.Font({ 'Family': "Verdana", 'PixelSize': 20 }) }),
				ui.VGap(20, 0),
				ui.ComboBox({'ID': 'Track', 'Text':'Video Track'}),
				ui.Button({'ID': 'ScanVersions', 'Text': 'Scan For Latest Versions', 'Weight':0}),
				ui.VGap(15, 0),
				ui.Label({'Text': 'Shots', 'Weight':0}),
				ui.Tree({'ID': 'ShotTree'}),
				ui.VGap(15, 0),
				ui.Label({'Text': 'Where to import new shots:', 'Weight':0}),
				ui.ComboBox({'ID': 'Location'}),
				ui.VGap(15, 0),
				ui.Button({'ID': 'Submit', 'Text': 'Update Versions!', 'Weight':0}),
				ui.Label({'Text': "Made by Austin ♥", 'Weight':0, 'Font':ui.Font({'PixelSize': 10})})
			])
		)

		self.window.Find('Location').AddItems(['Currently Open Bin', 'Same Bin As Original Clip'])


		# Register Events
		self.window.On[self.winID].Close = self.closeEvent
		self.window.On['ScanVersions'].Clicked = self.scanVersions
		self.window.On['Submit'].Clicked = self.versionUp

	def versionUp(self, event):
		if not self.shots:
			return

		tree = self.window.Find('ShotTree')

		importToSourceBin = self.window.Find('Location').CurrentText == "Same Bin As Original Clip"

		i = 0
		for shot in self.shots:
			success = shot.update(importToSourceBin)
			row = tree.TopLevelItem(i)
			if success:
				row.Text[1] = shot.highestVersion
			else:
				row.Text[0] = '!! FAILED !! ' + row.Text[0]
			i+=1

	def populateTrackList(self):
		dropdown = self.window.Find('Track')
		timeline = self.project.GetCurrentTimeline()

		videoTracks = [timeline.GetTrackName('video', i+1) for i in range(timeline.GetTrackCount('video'))]
		videoTracks.append('All Tracks')
		videoTracks.reverse()

		dropdown.AddItems(videoTracks)

	def scanVersions(self, event):
		timeline = self.project.GetCurrentTimeline()

		targetTrack = self.window.Find('Track').CurrentText
		
		# Grab all clips in selected track
		clips = []
		if targetTrack == 'All Tracks':
			for track in [(i+1, timeline.GetTrackName('video', i+1)) for i in range(timeline.GetTrackCount('video'))]:
				index, name = track
				clips += timeline.GetItemListInTrack('video', index)

		else:
			for track in [(i+1, timeline.GetTrackName('video', i+1)) for i in range(timeline.GetTrackCount('video'))]:
				index, name = track
				if name == targetTrack:
					clips += timeline.GetItemListInTrack('video', index)
					break

		self.shots = [Shot(i) for i in clips]
		self.shots = [i for i in self.shots if i.isVersionable]

		self.buildShotList()

		bad = [i for i in self.shots if i.highestVersion != i.highestInvalidVersion]

		if len(bad) > 0:
			self.alert(bad)


	def alert(self, shots):
		alert = self.dispatcher.AddWindow({
			'ID': self.winID + 'Error',
			'Geometry': [ 100,100,400,150 ],
			'WindowTitle': "Warning!",
			},
			ui.VGroup([
				ui.Label({ 'Text': "Warning!", 'Weight':0, 'Font': ui.Font({ 'Family': "Verdana", 'PixelSize': 20 }) }),
				ui.VGap(20, 0),
				ui.Label({'Text': 'Some shots had latest versions that were either \nmissing frames or not long enough. \nUsing lower versions on these shots for now.','WordWrap':True, 'Weight':0}),
			])
		)
		alert.On[self.winID + 'Error'].Close = lambda ev: alert.Hide()

		alert.Show()


	def buildShotList(self):
		tree = self.window.Find('ShotTree')
		# Reset list
		tree.Clear()

		header = tree.NewItem()
		header.Text[0] = 'Shot'
		header.Text[1] = 'Current'
		header.Text[2] = 'Highest'
		header.Text[3] = 'Highest Valid Render'
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
				row.Text[1] = shot.currentVersion
				row.Text[2] = shot.highestInvalidVersion
				row.Text[3] = shot.highestVersion
				tree.AddTopLevelItem(row)

	def closeEvent(self, event):
		self.dispatcher.ExitLoop()



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
		global versionRegex
		self.name = trackItem.GetName()
		self.trackItem = trackItem
		self.duration = trackItem.GetDuration()
		self.invalidVersions = []

		self.mpItem = trackItem.GetMediaPoolItem()
		self.path = self.mpItem.GetClipProperty('File Path')

		# restrict to unique entries. Proper version paths should be left with 1 unique version
		version = list(set(versionRegex.findall(self.path)))

		if len(version) == 0:
			self.isVersionable = False

		self.isVersionable = True
		self.currentVersion = version[-1]

		available = self.availableVersions()

		self.highestVersion = available[-1]

	def availableVersions(self):
		global versionRegex, frameSequenceRegex

		version = list(set(versionRegex.findall(self.path)))[-1]
		globPath = self.path.replace(version, '*')

		frameRange = frameSequenceRegex.findall(globPath)
		self.isSequence = isSequence = len(frameRange) > 0
		globPath = re.sub(frameSequenceRegex, '*', globPath)

		results = sorted(list(set([versionRegex.findall(i)[-1] for i in sorted(glob.glob(globPath))])))

		validResults = self.validateVersions(results)

		self.invalidVersions = [i for i in results if i not in validResults]

		self.highestInvalidVersion = results[-1] 

		return validResults

	def validateVersions(self, versions):
		if not self.isSequence:
			return versions

		goodVersions = []

		for version in versions:
			versionPath = self.path.replace(self.currentVersion, version)
			sequence = frameSequenceRegex.findall(versionPath)[-1]
			globPath = versionPath.replace(sequence, '*')

			files = glob.glob(globPath)
			frames = sorted([i.replace(globPath.split('*')[0], '').replace(globPath.split('*')[1], '') for i in files])
			
			if not self.missingFrames(frames) and len(frames) >= self.duration:
				goodVersions.append(version)

		return goodVersions

	def missingFrames(self, frames):
		start = int(frames[0])
		i = start
		for frame in frames:
			if i == int(frame):
				i+=1
			else:
				return True
		return False

	def update(self, importToSourceBin=False):
		global resolve, project
		ms = resolve.GetMediaStorage()

		if self.currentVersion == self.highestVersion:
			return True

		newPath = self.path.replace(self.currentVersion, self.highestVersion)
		if self.isSequence:
			newPath = newPath.replace(re.findall(r'[\\\/]{1}([^\\\/]+)(?!.+)', newPath)[-1], '')


		if importToSourceBin:
			folder = self.findFolder(self.mpItem)
			project.GetMediaPool().SetCurrentFolder(folder)

		ms.AddItemListToMediaPool(newPath)

		item = self.findItemInProject(newPath)

		if not item:
			return False

		self.swap(item)

		if newPath in self.trackItem.GetMediaPoolItem().GetClipProperty('File Path'):
			return True

		return False


	def swap(self, mediaPoolItem):
		timelineItem = self.trackItem

		leftOffset = timelineItem.GetLeftOffset()
		start = int(timelineItem.GetMediaPoolItem().GetClipProperty('Start'))
		newIn = start + leftOffset

		timelineItem.AddTake(mediaPoolItem, newIn, newIn + self.duration)
		timelineItem.SelectTakeByIndex(timelineItem.GetTakesCount())
		timelineItem.FinalizeTake()


	def findItemInProject(self, path, folder=None):
		global project
		if not folder:
			folder = project.GetMediaPool().GetRootFolder()

		matches = [i for i in folder.GetClipList() if path in i.GetClipProperty('File Path')]
		if len(matches) > 0:
			return matches[-1]

		for subfolder in folder.GetSubFolderList():
			result = self.findItemInProject(path, subfolder)
			if result:
				return result

		return False

	def findFolder(self, item, folder=None):
		global project
		if not folder:
			folder = project.GetMediaPool().GetRootFolder()

		matches = [i for i in folder.GetClipList() if i.GetClipProperty() == item.GetClipProperty()]
		if len(matches) > 0:
			return folder

		for subfolder in folder.GetSubFolderList():
			result = self.findFolder(item, subfolder)
			if result:
				return result

		return False






VersionUpShots()