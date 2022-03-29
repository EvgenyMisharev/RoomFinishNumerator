﻿using Autodesk.Revit.DB;
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

                        List<FamilyInstance> doorsInRoomList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Doors)
                            .OfClass(typeof(FamilyInstance))
                            .WhereElementIsNotElementType()
                            .Cast<FamilyInstance>()
                            .Where(d => d.Room != null)
                            .Where(d => d.Room.Id == room.Id)
                            .ToList();

                        List<FamilyInstance> doorsFromRoomList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Doors)
                            .OfClass(typeof(FamilyInstance))
                            .WhereElementIsNotElementType()
                            .Cast<FamilyInstance>()
                            .Where(d => d.FromRoom != null)
                            .Where(d => d.FromRoom.Id == room.Id)
                            .ToList();

                        foreach (FamilyInstance door in doorsFromRoomList)
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

                            doorsInRoomArea += maxWidth * maxWidth;
                        }

                        List<FamilyInstance> windowsInRoomList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Windows)
                            .OfClass(typeof(FamilyInstance))
                            .WhereElementIsNotElementType()
                            .Cast<FamilyInstance>()
                            .Where(d => d.Room != null)
                            .Where(d => d.Room.Id == room.Id)
                            .ToList();

                        List<FamilyInstance> windowsFromRoomList = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Windows)
                            .OfClass(typeof(FamilyInstance))
                            .WhereElementIsNotElementType()
                            .Cast<FamilyInstance>()
                            .Where(d => d.FromRoom != null)
                            .Where(d => d.FromRoom.Id == room.Id)
                            .ToList();

                        foreach (FamilyInstance window in windowsFromRoomList)
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

                            windowssInRoomArea += maxWidth * maxWidth;
                        }

                        //СПОРНАЯ ИДЕЯ С ВИТРАЖАМИ
                        //SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions();
                        //IList<IList<BoundarySegment>> boundarySegmentsList = room.GetBoundarySegments(opt);
                        //foreach(IList<BoundarySegment> lbs in boundarySegmentsList)
                        //{
                        //    foreach(BoundarySegment bs in lbs)
                        //    {
                        //        Element boundaryElement = doc.GetElement(bs.ElementId);
                        //        Wall boundaryWall = null;
                        //        try
                        //        {
                        //            boundaryWall = boundaryElement as Wall;
                        //        }
                        //        catch 
                        //        { 

                        //        }
                        //        if (boundaryWall != null)
                        //        {
                        //            if(boundaryWall.CurtainGrid != null)
                        //            {
                        //                curtainWallArea += boundaryWall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                        //            }
                        //        }
                        //    }
                        //}

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