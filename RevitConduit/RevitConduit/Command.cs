#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

#endregion

namespace RevitConduit
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        private const string ELEMENT_NAME = "CONDUIT_PATH_POINT";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            FilteredElementCollector collectorType = new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.Electrical.ConduitType));
            FilteredElementCollector collectorLevel = new FilteredElementCollector(doc).OfClass(typeof(Level));

            Autodesk.Revit.DB.Electrical.ConduitType type = collectorType.FirstElement() as Autodesk.Revit.DB.Electrical.ConduitType;
            Level level = collectorLevel.FirstElement() as Level;

            List<Element> listOfElements = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).WhereElementIsElementType().ToElements().Where(e => e.Name == ELEMENT_NAME).ToList<Element>();

            ElementId elementId = listOfElements[0].Id;

            FamilyInstanceFilter familyInstanceFilter = new FamilyInstanceFilter(doc, elementId);

            IList<Element> familyInstances = new FilteredElementCollector(doc).WherePasses(familyInstanceFilter).ToElements();

            List<XYZ> listLocation = new List<XYZ>();

            List<Element> elementsConduit = new List<Element>();

            int step = 0;

            foreach (var conduitPathPoint in familyInstances)
            {
                LocationPoint locationPoint = conduitPathPoint.Location as LocationPoint;

                var location = locationPoint.Point;
                listLocation.Add(location);

                if (step >= 1)
                {
                    Transaction trans = new Transaction(doc);

                    trans.Start("createConduit");

                    var coduit = Autodesk.Revit.DB.Electrical.Conduit.Create(doc, type.Id, listLocation[0], listLocation[1], level.Id);

                    trans.Commit();

                    elementsConduit.Add(coduit);
                    if (step >= 2 && step <= 4)
                    {
                        trans = new Transaction(doc);

                        trans.Start("createConnect");
                        Connect(listLocation[0], elementsConduit[0], elementsConduit[1], doc);
                        trans.Commit();
                    }

                    listLocation.Clear();
                    elementsConduit.Clear();
                    listLocation.Add(location);
                    elementsConduit.Add(coduit);
                }
                step++;
            }
            return Result.Succeeded;
        }
        private static void Connect(XYZ location, Element a, Element b, Document doc)
        {

            ConnectorManager cm = GetConnectorManager(a);

            if (null == cm)
            {
                throw new ArgumentException(
                  "Element a has no connectors.");
            }

            Connector ca = GetConnectorClosestTo(cm.Connectors, location);

            cm = GetConnectorManager(b);

            if (null == cm)
            {
                throw new ArgumentException(
                  "Element b has no connectors.");
            }

            Connector cb = GetConnectorClosestTo(cm.Connectors, location);
            doc.Create.NewElbowFitting(ca, cb);
        }

        private static ConnectorManager GetConnectorManager(Element e)
        {
            MEPCurve mc = e as MEPCurve;
            FamilyInstance fi = e as FamilyInstance;

            if (null == mc && null == fi)
            {
                throw new ArgumentException(
                  "Element is neither an MEP curve nor a fitting.");
            }

            return null == mc
              ? fi.MEPModel.ConnectorManager
              : mc.ConnectorManager;
        }

        private static Connector GetConnectorClosestTo(ConnectorSet connectors, XYZ location)
        {
            Connector targetConnector = null;
            double minDist = double.MaxValue;

            foreach (Connector c in connectors)
            {
                double d = c.Origin.DistanceTo(location);

                if (d < minDist)
                {
                    targetConnector = c;
                    minDist = d;
                }
            }
            return targetConnector;
        }
    }
}
