using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RoomFinishNumerator
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class RoomFinishNumeratorCommand : IExternalCommand
    {
        RoomFinishNumeratorProgressBarWPF roomFinishNumeratorProgressBarWPF;
        RoomFinishNumeratorOpeningsProgressBarWPF roomFinishNumeratorOpeningsProgressBarWPF;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            RoomFinishNumeratorWPF roomFinishNumeratorWPF = new RoomFinishNumeratorWPF();
            roomFinishNumeratorWPF.ShowDialog();
            if (roomFinishNumeratorWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }

            string roomFinishNumberingSelectedName = roomFinishNumeratorWPF.RoomFinishNumberingSelectedName;
            bool considerCeilings = roomFinishNumeratorWPF.ConsiderCeilings;
            bool considerOpenings = roomFinishNumeratorWPF.ConsiderOpenings;


            using (Transaction t = new Transaction(doc))
            {
                t.Start("Вычисление площадей проемов");
                if (considerOpenings)
                {
                    List<Room> roomList = new FilteredElementCollector(doc)
                        .OfClass(typeof(SpatialElement))
                        .WhereElementIsNotElementType()
                        .Where(r => r.GetType() == typeof(Room))
                        .Cast<Room>()
                        .Where(r => r.Area > 0)
                        .OrderBy(r => (doc.GetElement(r.LevelId) as Level).Elevation)
                        .ToList();

                    Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPointForOpenings));
                    newWindowThread.SetApartmentState(ApartmentState.STA);
                    newWindowThread.IsBackground = true;
                    newWindowThread.Start();
                    int step = 0;
                    Thread.Sleep(100);
                    roomFinishNumeratorOpeningsProgressBarWPF.pb_RoomFinishNumeratorOpeningsProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorOpeningsProgressBarWPF.pb_RoomFinishNumeratorOpeningsProgressBar.Minimum = 0);
                    roomFinishNumeratorOpeningsProgressBarWPF.pb_RoomFinishNumeratorOpeningsProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorOpeningsProgressBarWPF.pb_RoomFinishNumeratorOpeningsProgressBar.Maximum = roomList.Count);

                    foreach (Room room in roomList)
                    {
                        step++;
                        roomFinishNumeratorOpeningsProgressBarWPF.pb_RoomFinishNumeratorOpeningsProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorOpeningsProgressBarWPF.pb_RoomFinishNumeratorOpeningsProgressBar.Value = step);

                        double doorsInRoomArea = 0;
                        double windowssInRoomArea = 0;
                        double curtainWallArea = 0;

                        List<FamilyInstance> doorsOnRoomLevelList = new FilteredElementCollector(doc)
                           .OfCategory(BuiltInCategory.OST_Doors)
                           .OfClass(typeof(FamilyInstance))
                           .WhereElementIsNotElementType()
                           .Cast<FamilyInstance>()
                           .Where(d => d.LevelId == room.LevelId)
                           .ToList();


                        List<FamilyInstance> doorsInRoomList = doorsOnRoomLevelList
                            .Where(d => d.Room != null)
                            .Where(d => d.Room.Id == room.Id)
                            .ToList();

                        List<FamilyInstance> doorsFromRoomList = doorsOnRoomLevelList
                            .Where(d => d.FromRoom != null)
                            .Where(d => d.FromRoom.Id == room.Id)
                            .ToList();

                        List<FamilyInstance> doorsToRoomList = doorsOnRoomLevelList
                            .Where(d => d.ToRoom != null)
                            .Where(d => d.ToRoom.Id == room.Id)
                            .ToList();

                        foreach (FamilyInstance door in doorsFromRoomList)
                        {
                            if (!doorsInRoomList.Contains(door))
                            {
                                doorsInRoomList.Add(door);
                            }
                        }

                        foreach (FamilyInstance door in doorsToRoomList)
                        {
                            if (!doorsInRoomList.Contains(door))
                            {
                                doorsInRoomList.Add(door);
                            }
                        }

                        foreach (FamilyInstance door in doorsInRoomList)
                        {
                            double roughHeight = 0;
                            double roughWidth = 0;
                            double caseworkHeight = 0;
                            double caseworkWidth = 0;
                            double maxHeight = 0;
                            double maxWidth = 0;

                            if (door.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM) != null)
                            {
                                roughHeight = door.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM).AsDouble();
                            }
                            if (door.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM) != null)
                            {
                                roughWidth = door.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM).AsDouble();
                            }
                            if (door.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT) != null)
                            {
                                caseworkHeight = door.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT).AsDouble();
                            }
                            if (door.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH) != null)
                            {
                                caseworkWidth = door.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH).AsDouble();
                            }

                            if (roughHeight >= caseworkHeight)
                            {
                                maxHeight = roughHeight;
                            }
                            else
                            {
                                maxHeight = caseworkHeight;
                            }

                            if (roughWidth >= caseworkWidth)
                            {
                                maxWidth = roughWidth;
                            }
                            else
                            {
                                maxWidth = caseworkWidth;
                            }

                            doorsInRoomArea += maxHeight * maxWidth;
                        }

                        List<FamilyInstance> windowsOnRoomLevelList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Windows)
                            .OfClass(typeof(FamilyInstance))
                            .WhereElementIsNotElementType()
                            .Cast<FamilyInstance>()
                            .Where(w => w.LevelId == room.LevelId)
                            .ToList();

                        List<FamilyInstance> windowsInRoomList = windowsOnRoomLevelList
                            .Where(w => w.Room != null)
                            .Where(w => w.Room.Id == room.Id)
                            .ToList();

                        List<FamilyInstance> windowsFromRoomList = windowsOnRoomLevelList
                            .Where(w => w.FromRoom != null)
                            .Where(w => w.FromRoom.Id == room.Id)
                            .ToList();

                        List<FamilyInstance> windowsToRoomList = windowsOnRoomLevelList
                            .Where(w => w.ToRoom != null)
                            .Where(w => w.ToRoom.Id == room.Id)
                            .ToList();

                        foreach (FamilyInstance window in windowsFromRoomList)
                        {
                            if (!windowsInRoomList.Contains(window))
                            {
                                windowsInRoomList.Add(window);
                            }
                        }

                        foreach (FamilyInstance window in windowsToRoomList)
                        {
                            if (!windowsInRoomList.Contains(window))
                            {
                                windowsInRoomList.Add(window);
                            }
                        }

                        foreach (FamilyInstance window in windowsInRoomList)
                        {
                            double roughHeight = 0;
                            double roughWidth = 0;
                            double caseworkHeight = 0;
                            double caseworkWidth = 0;
                            double maxHeight = 0;
                            double maxWidth = 0;

                            if (window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM) != null)
                            {
                                roughHeight = window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM).AsDouble();
                            }
                            if (window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM) != null)
                            {
                                roughWidth = window.Symbol.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM).AsDouble();
                            }
                            if (window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT) != null)
                            {
                                caseworkHeight = window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_HEIGHT).AsDouble();
                            }
                            if (window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH) != null)
                            {
                                caseworkWidth = window.Symbol.get_Parameter(BuiltInParameter.CASEWORK_WIDTH).AsDouble();
                            }

                            if (roughHeight >= caseworkHeight)
                            {
                                maxHeight = roughHeight;
                            }
                            else
                            {
                                maxHeight = caseworkHeight;
                            }

                            if (roughWidth >= caseworkWidth)
                            {
                                maxWidth = roughWidth;
                            }
                            else
                            {
                                maxWidth = caseworkWidth;
                            }

                            windowssInRoomArea += maxHeight * maxWidth;
                        }

                        Solid roomSolid = null;
                        GeometryElement geomRoomElement = room.get_Geometry(new Options());
                        foreach (GeometryObject geomObj in geomRoomElement)
                        {
                            roomSolid = geomObj as Solid;
                            if (roomSolid != null) break;
                        }
                        if(roomSolid != null)
                        {
                            List<Wall> curtainWallsList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .OfClass(typeof(Wall))
                                .WhereElementIsNotElementType()
                                .Cast<Wall>()
                                .Where(w => w.LevelId == room.LevelId)
                                .Where(w => w.CurtainGrid != null)
                                .ToList();

                            SolidCurveIntersectionOptions intersectOptions = new SolidCurveIntersectionOptions();
                            foreach (Wall wall in curtainWallsList)
                            {
                                List<ElementId> CurtainPanelsIdList = wall.CurtainGrid.GetPanelIds().ToList();
                                foreach (ElementId panelId in CurtainPanelsIdList)
                                {
                                    Panel panel = null;
                                    FamilyInstance doorwindows = null;
                                    panel = doc.GetElement(panelId) as Panel;
                                    if(panel == null)
                                    {
                                        doorwindows = doc.GetElement(panelId) as FamilyInstance;
                                        if(doorwindows != null)
                                        {
                                            double curtainWallPanelsHeight = 0;
                                            double curtainWallPanelsWidth = 0;

                                            if (doorwindows.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_HEIGHT) != null)
                                            {
                                                curtainWallPanelsHeight = doorwindows.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_HEIGHT).AsDouble();
                                            }
                                            if (doorwindows.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_WIDTH) != null)
                                            {
                                                curtainWallPanelsWidth = doorwindows.get_Parameter(BuiltInParameter.CURTAIN_WALL_PANELS_WIDTH).AsDouble();
                                            }
                                           
                                            curtainWallArea += curtainWallPanelsHeight * curtainWallPanelsWidth;
                                        }
                                    }

                                    if (panel != null)
                                    {
                                        BoundingBoxXYZ panelBoundingBox = panel.get_BoundingBox(null);
                                        if (panelBoundingBox == null) continue;
                                        XYZ panelCenter = (panelBoundingBox.Max + panelBoundingBox.Min) / 2;
                                        Curve lineA = Line.CreateBound(panelCenter, panelCenter + (600 / 304.8) * panel.FacingOrientation.Normalize()) as Curve;
                                        Curve lineB = Line.CreateBound(panelCenter, panelCenter + (600 / 304.8) * panel.FacingOrientation.Normalize().Negate()) as Curve;

                                        SolidCurveIntersection intersectionA = roomSolid.IntersectWithCurve(lineA, intersectOptions);
                                        SolidCurveIntersection intersectionB = roomSolid.IntersectWithCurve(lineB, intersectOptions);
                                        if(intersectionA.SegmentCount > 0 || intersectionB.SegmentCount > 0)
                                        {
                                            curtainWallArea += panel.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                                        }
                                    }
                                }
                            }
                        }

                        Guid openingsAreaParamGuid = new Guid("18e3f49d-1315-415f-8359-8f045a7a8938");
                        if (room.get_Parameter(openingsAreaParamGuid) != null)
                        {
                            room.get_Parameter(openingsAreaParamGuid).Set(doorsInRoomArea + windowssInRoomArea + curtainWallArea);
                        }
                    }
                    roomFinishNumeratorOpeningsProgressBarWPF.Dispatcher.Invoke(() => roomFinishNumeratorOpeningsProgressBarWPF.Close());
                }
                else
                {
                    List<Room> roomList = new FilteredElementCollector(doc)
                        .OfClass(typeof(SpatialElement))
                        .WhereElementIsNotElementType()
                        .Where(r => r.GetType() == typeof(Room))
                        .Cast<Room>()
                        .Where(r => r.Area > 0)
                        .OrderBy(r => (doc.GetElement(r.LevelId) as Level).Elevation)
                        .ToList();

                    foreach (Room room in roomList)
                    {
                        Guid openingsAreaParamGuid = new Guid("18e3f49d-1315-415f-8359-8f045a7a8938");
                        if (room.get_Parameter(openingsAreaParamGuid) != null)
                        {
                            room.get_Parameter(openingsAreaParamGuid).Set(0);
                        }
                    }
                }
                t.Commit();
            }

            if (roomFinishNumberingSelectedName == "rbt_EndToEndThroughoutTheProject")
            {
                List<Room> roomList = new FilteredElementCollector(doc)
                    .OfClass(typeof(SpatialElement))
                    .WhereElementIsNotElementType()
                    .Where(r => r.GetType() == typeof(Room))
                    .Cast<Room>()
                    .Where(r => r.Area > 0)
                    .OrderBy(r => (doc.GetElement(r.LevelId) as Level).Elevation)
                    .ToList();

                using (TransactionGroup tg = new TransactionGroup(doc))
                {
                    tg.Start("Нумерация отделки");
                    if (considerCeilings)
                    {
                        List<RoomProperties> roomPropertiesList = new List<RoomProperties>();
                        foreach (Room room in roomList)
                        {
                            RoomProperties tempRoomProperties = new RoomProperties();
                            tempRoomProperties.CeilingFinishStrParam = room.get_Parameter(BuiltInParameter.ROOM_FINISH_CEILING).AsString();
                            tempRoomProperties.WallFinishStrParam = room.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).AsString();
                            tempRoomProperties.BottomWallFinishStrParam = room.LookupParameter("АР_ОтделкаСтенСнизу").AsString();

                            //НЕ УВЕРЕН В ПРОВЕРКЕ!!!
                            if (!roomPropertiesList.Contains(tempRoomProperties))
                            {
                                roomPropertiesList.Add(tempRoomProperties);
                            }
                        }

                        Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                        newWindowThread.SetApartmentState(ApartmentState.STA);
                        newWindowThread.IsBackground = true;
                        newWindowThread.Start();
                        int step = 0;
                        Thread.Sleep(100);
                        roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Minimum = 0);
                        roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Maximum = roomPropertiesList.Count);

                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Внесение номеров в помещения");
                            foreach (RoomProperties rp in roomPropertiesList)
                            {
                                step++;
                                roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Value = step);

                                List<Room> roomListForNumbering = new FilteredElementCollector(doc)
                                    .OfClass(typeof(SpatialElement))
                                    .WhereElementIsNotElementType()
                                    .Where(r => r.GetType() == typeof(Room))
                                    .Cast<Room>()
                                    .Where(r => r.Area > 0)
                                    .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_CEILING) != null)
                                    .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_CEILING).AsString() == rp.CeilingFinishStrParam)
                                    .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL) != null)
                                    .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).AsString() == rp.WallFinishStrParam)
                                    .Where(r => r.LookupParameter("АР_ОтделкаСтенСнизу") != null)
                                    .Where(r => r.LookupParameter("АР_ОтделкаСтенСнизу").AsString() == rp.BottomWallFinishStrParam)
                                    .OrderBy(r => r.Number, new AlphanumComparatorFastString())
                                    .ToList();

                                string roomNumbersByRoom = null;
                                foreach (Room r in roomListForNumbering)
                                {
                                    if (roomNumbersByRoom == null)
                                    {
                                        roomNumbersByRoom += r.Number;
                                    }
                                    else
                                    {
                                        roomNumbersByRoom += ", " + r.Number;
                                    }
                                }


                                foreach (Room r in roomListForNumbering)
                                {
                                    r.LookupParameter("АР_НомераПомещенийВедОтделки").Set(roomNumbersByRoom);
                                }
                            }
                            roomFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.Close());
                            t.Commit();
                        }
                    }
                    else
                    {
                        List<RoomProperties> roomPropertiesList = new List<RoomProperties>();
                        foreach (Room room in roomList)
                        {
                            RoomProperties tempRoomProperties = new RoomProperties();
                            tempRoomProperties.WallFinishStrParam = room.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).AsString();
                            tempRoomProperties.BottomWallFinishStrParam = room.LookupParameter("АР_ОтделкаСтенСнизу").AsString();

                            //НЕ УВЕРЕН В ПРОВЕРКЕ!!!
                            if (!roomPropertiesList.Contains(tempRoomProperties))
                            {
                                roomPropertiesList.Add(tempRoomProperties);
                            }
                        }

                        Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                        newWindowThread.SetApartmentState(ApartmentState.STA);
                        newWindowThread.IsBackground = true;
                        newWindowThread.Start();
                        int step = 0;
                        Thread.Sleep(100);
                        roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Minimum = 0);
                        roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Maximum = roomPropertiesList.Count);

                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Внесение номеров в помещения");
                            foreach (RoomProperties rp in roomPropertiesList)
                            {
                                step++;
                                roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Value = step);

                                List<Room> roomListForNumbering = new FilteredElementCollector(doc)
                                    .OfClass(typeof(SpatialElement))
                                    .WhereElementIsNotElementType()
                                    .Where(r => r.GetType() == typeof(Room))
                                    .Cast<Room>()
                                    .Where(r => r.Area > 0)
                                    .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL) != null)
                                    .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).AsString() == rp.WallFinishStrParam)
                                    .Where(r => r.LookupParameter("АР_ОтделкаСтенСнизу") != null)
                                    .Where(r => r.LookupParameter("АР_ОтделкаСтенСнизу").AsString() == rp.BottomWallFinishStrParam)
                                    .OrderBy(r => r.Number, new AlphanumComparatorFastString())
                                    .ToList();

                                string roomNumbersByRoom = null;
                                foreach (Room r in roomListForNumbering)
                                {
                                    if (roomNumbersByRoom == null)
                                    {
                                        roomNumbersByRoom += r.Number;
                                    }
                                    else
                                    {
                                        roomNumbersByRoom += ", " + r.Number;
                                    }
                                }


                                foreach (Room r in roomListForNumbering)
                                {
                                    r.LookupParameter("АР_НомераПомещенийВедОтделки").Set(roomNumbersByRoom);
                                }
                            }
                            roomFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.Close());
                            t.Commit();
                        }
                    }
                    tg.Assimilate();
                }

            }
            else if (roomFinishNumberingSelectedName == "rbt_SeparatedByLevels")
            {
                List<Level> levelList = new FilteredElementCollector(doc)
                   .OfClass(typeof(Level))
                   .WhereElementIsNotElementType()
                   .Cast<Level>()
                   .OrderBy(l => l.Elevation)
                   .ToList();

                Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                int step = 0;
                Thread.Sleep(100);
                roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Minimum = 0);
                roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Maximum = levelList.Count);

                using (TransactionGroup tg = new TransactionGroup(doc))
                {
                    tg.Start("Нумерация отделки");
                    if (considerCeilings)
                    {
                        using (Transaction t = new Transaction(doc))
                        {
                            t.Start("Внесение номеров в помещения");
                            foreach (Level lv in levelList)
                            {
                                step++;
                                roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Value = step);

                                List<Room> roomList = new FilteredElementCollector(doc)
                                    .OfClass(typeof(SpatialElement))
                                    .WhereElementIsNotElementType()
                                    .Where(r => r.GetType() == typeof(Room))
                                    .Cast<Room>()
                                    .Where(r => r.Area > 0)
                                    .Where(r => r.LevelId == lv.Id)
                                    .ToList();

                                List<RoomProperties> roomPropertiesList = new List<RoomProperties>();
                                foreach (Room room in roomList)
                                {
                                    RoomProperties tempRoomProperties = new RoomProperties();
                                    tempRoomProperties.CeilingFinishStrParam = room.get_Parameter(BuiltInParameter.ROOM_FINISH_CEILING).AsString();
                                    tempRoomProperties.WallFinishStrParam = room.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).AsString();
                                    tempRoomProperties.BottomWallFinishStrParam = room.LookupParameter("АР_ОтделкаСтенСнизу").AsString();

                                    //НЕ УВЕРЕН В ПРОВЕРКЕ!!!
                                    if (!roomPropertiesList.Contains(tempRoomProperties))
                                    {
                                        roomPropertiesList.Add(tempRoomProperties);
                                    }
                                }
                                foreach (RoomProperties rp in roomPropertiesList)
                                {
                                    List<Room> roomListForNumbering = new FilteredElementCollector(doc)
                                        .OfClass(typeof(SpatialElement))
                                        .WhereElementIsNotElementType()
                                        .Where(r => r.GetType() == typeof(Room))
                                        .Cast<Room>()
                                        .Where(r => r.Area > 0)
                                        .Where(r => r.LevelId == lv.Id)
                                        .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_CEILING) != null)
                                        .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_CEILING).AsString() == rp.CeilingFinishStrParam)
                                        .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL) != null)
                                        .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).AsString() == rp.WallFinishStrParam)
                                        .Where(r => r.LookupParameter("АР_ОтделкаСтенСнизу") != null)
                                        .Where(r => r.LookupParameter("АР_ОтделкаСтенСнизу").AsString() == rp.BottomWallFinishStrParam)
                                        .OrderBy(r => r.Number, new AlphanumComparatorFastString())
                                        .ToList();

                                    string roomNumbersByRoom = null;
                                    foreach (Room r in roomListForNumbering)
                                    {
                                        if (roomNumbersByRoom == null)
                                        {
                                            roomNumbersByRoom += r.Number;
                                        }
                                        else
                                        {
                                            roomNumbersByRoom += ", " + r.Number;
                                        }
                                    }


                                    foreach (Room r in roomListForNumbering)
                                    {
                                        r.LookupParameter("АР_НомераПомещенийВедОтделки").Set(roomNumbersByRoom);
                                    }
                                }
                            }
                            roomFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.Close());
                            t.Commit();
                        }

                    }
                    else
                    {
                        using (Transaction t = new Transaction(doc)) 
                        {
                            t.Start("Внесение номеров в помещения");
                            foreach (Level lv in levelList)
                            {
                                step++;
                                roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.pb_RoomFinishNumeratorProgressBar.Value = step);

                                List<Room> roomList = new FilteredElementCollector(doc)
                                    .OfClass(typeof(SpatialElement))
                                    .WhereElementIsNotElementType()
                                    .Where(r => r.GetType() == typeof(Room))
                                    .Cast<Room>()
                                    .Where(r => r.Area > 0)
                                    .Where(r => r.LevelId == lv.Id)
                                    .ToList();

                                List<RoomProperties> roomPropertiesList = new List<RoomProperties>();
                                foreach (Room room in roomList)
                                {
                                    RoomProperties tempRoomProperties = new RoomProperties();
                                    tempRoomProperties.WallFinishStrParam = room.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).AsString();
                                    tempRoomProperties.BottomWallFinishStrParam = room.LookupParameter("АР_ОтделкаСтенСнизу").AsString();

                                    //НЕ УВЕРЕН В ПРОВЕРКЕ!!!
                                    if (!roomPropertiesList.Contains(tempRoomProperties))
                                    {
                                        roomPropertiesList.Add(tempRoomProperties);
                                    }
                                }
                                foreach (RoomProperties rp in roomPropertiesList)
                                {
                                    List<Room> roomListForNumbering = new FilteredElementCollector(doc)
                                        .OfClass(typeof(SpatialElement))
                                        .WhereElementIsNotElementType()
                                        .Where(r => r.GetType() == typeof(Room))
                                        .Cast<Room>()
                                        .Where(r => r.Area > 0)
                                        .Where(r => r.LevelId == lv.Id)
                                        .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL) != null)
                                        .Where(r => r.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).AsString() == rp.WallFinishStrParam)
                                        .Where(r => r.LookupParameter("АР_ОтделкаСтенСнизу") != null)
                                        .Where(r => r.LookupParameter("АР_ОтделкаСтенСнизу").AsString() == rp.BottomWallFinishStrParam)
                                        .OrderBy(r => r.Number, new AlphanumComparatorFastString())
                                        .ToList();

                                    string roomNumbersByRoom = null;
                                    foreach (Room r in roomListForNumbering)
                                    {
                                        if (roomNumbersByRoom == null)
                                        {
                                            roomNumbersByRoom += r.Number;
                                        }
                                        else
                                        {
                                            roomNumbersByRoom += ", " + r.Number;
                                        }
                                    }


                                    foreach (Room r in roomListForNumbering)
                                    {
                                        r.LookupParameter("АР_НомераПомещенийВедОтделки").Set(roomNumbersByRoom);
                                    }
                                }
                            }
                            roomFinishNumeratorProgressBarWPF.Dispatcher.Invoke(() => roomFinishNumeratorProgressBarWPF.Close());
                            t.Commit();
                        }
                    }
                    tg.Assimilate();
                }

            }
            return Result.Succeeded;
        }
        private void ThreadStartingPoint()
        {
            roomFinishNumeratorProgressBarWPF = new RoomFinishNumeratorProgressBarWPF();
            roomFinishNumeratorProgressBarWPF.Show();
            System.Windows.Threading.Dispatcher.Run();
        }
        private void ThreadStartingPointForOpenings()
        {
            roomFinishNumeratorOpeningsProgressBarWPF = new RoomFinishNumeratorOpeningsProgressBarWPF();
            roomFinishNumeratorOpeningsProgressBarWPF.Show();
            System.Windows.Threading.Dispatcher.Run();
        }
    }
}
