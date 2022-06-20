using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;


namespace AROTask_2
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //все помещения из документа
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            RoomFilter filter = new RoomFilter();
            IList<Element> rooms = collector.WherePasses(filter).ToElements();


            foreach (Room room in rooms)
            {
                //проверка наличия коэффициента
                Parameter coefParameter = room.LookupParameter("ARO_Коэффициент площади");
                double coefValue;
                if (coefParameter.AsString() != null)
                {
                    coefValue = Double.Parse(coefParameter.AsString(), System.Globalization.CultureInfo.InvariantCulture);
                }
                else
                {
                    coefValue = 1;
                }


                Parameter areaParameter = room.get_Parameter(BuiltInParameter.ROOM_AREA);
                double roomArea = UnitUtils.ConvertFromInternalUnits(areaParameter.AsDouble(), UnitTypeId.SquareMeters);

                //новая площадь в категорию "помещения"
                var categorySet = new CategorySet();
                categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Rooms));

                using (Transaction ts = new Transaction(doc, "set parameter"))
                {
                    ts.Start();
                    CreateSharedParameter(uiapp.Application, doc, "ARO_Площадь с коэффициентом", categorySet, BuiltInParameterGroup.PG_DATA, true);
                    Parameter dParameter = room.LookupParameter("ARO_Площадь с коэффициентом");
                    if (dParameter != null)
                    {
                        dParameter.Set($"{roomArea * coefValue:f2}");
                    }
                    ts.Commit();
                }
            }

            return Result.Succeeded;
        }

        //метод создает дополнительные параметры в документе из ФОП1.txt
        private void CreateSharedParameter(Application application,
            Document doc, string parameterName, CategorySet categorySet,
            BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile();
            if (definitionFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return;
            }

            Definition definition = definitionFile.Groups
                .SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals(parameterName));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap map = doc.ParameterBindings;
            map.Insert(definition, binding, builtInParameterGroup);
        }
    }
}
