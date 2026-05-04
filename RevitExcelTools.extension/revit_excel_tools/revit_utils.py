"""
Revit-specific utilities for data collection and updating
"""
from Autodesk.Revit.DB import FilteredElementCollector, ElementCategoryFilter, BuiltInCategory
from pyrevit import revit

class RevitDataCollector(object):
    """Collect element data from Revit document"""

    def __init__(self, doc):
        self.doc = doc

    def collect_elements(self, element_filter=None):
        """
        Collect elements from Revit document

        Args:
            element_filter: List of category names to filter (e.g., ["Walls", "Doors"])

        Returns:
            List of dictionaries containing element data
        """
        elements = []

        try:
            # Get all elements
            collector = FilteredElementCollector(self.doc)
            all_elements = collector.WhereElementIsNotElementType().ToElements()

            for elem in all_elements:
                try:
                    elem_data = self._extract_element_data(elem)
                    if elem_data:
                        elements.append(elem_data)
                except Exception as e:
                    print(f"Error processing element {elem.Id}: {e}")
                    continue

        except Exception as e:
            print(f"Error collecting elements: {e}")

        return elements

    def _extract_element_data(self, elem):
        """Extract relevant data from an element"""
        try:
            data = {
                "ElementId": elem.Id.IntegerValue,
                "Category": elem.Category.Name if elem.Category else "Unknown",
                "Name": elem.Name if hasattr(elem, 'Name') else "",
            }

            # Extract common parameters
            params = self._get_element_parameters(elem)
            data.update(params)

            return data
        except Exception as e:
            print(f"Error extracting data: {e}")
            return None

    def _get_element_parameters(self, elem):
        """Extract parameters from an element"""
        params = {}
        try:
            for param in elem.Parameters:
                try:
                    param_name = param.Definition.Name
                    param_value = self._get_parameter_value(param)
                    params[param_name] = param_value
                except Exception:
                    continue
        except Exception as e:
            print(f"Error extracting parameters: {e}")

        return params

    def _get_parameter_value(self, param):
        """Get the value of a parameter"""
        try:
            if param.HasValue:
                if param.StorageType.ToString() == "String":
                    return param.AsString()
                elif param.StorageType.ToString() == "Double":
                    return param.AsDouble()
                elif param.StorageType.ToString() == "Integer":
                    return param.AsInteger()
                else:
                    return param.AsValueString()
            return ""
        except Exception:
            return ""


class RevitDataUpdater(object):
    """Update Revit elements with data from external sources"""

    def __init__(self, doc):
        self.doc = doc

    def update_elements(self, data):
        """
        Update Revit elements with provided data

        Args:
            data: List of dictionaries with ElementId and parameters to update

        Returns:
            Dictionary with update results
        """
        results = {"updated": 0, "failed": 0, "skipped": 0}

        with revit.Transaction("Update Elements from Excel"):
            for item in data:
                try:
                    elem_id = item.get("ElementId")
                    if not elem_id:
                        results["skipped"] += 1
                        continue

                    elem = self.doc.GetElement(elem_id)
                    if not elem:
                        results["skipped"] += 1
                        continue

                    # Update parameters
                    params_updated = self._update_element_parameters(elem, item)
                    if params_updated:
                        results["updated"] += 1
                    else:
                        results["skipped"] += 1

                except Exception as e:
                    print(f"Error updating element: {e}")
                    results["failed"] += 1
                    continue

        return results

    def _update_element_parameters(self, elem, data):
        """Update parameters of an element"""
        updated = False
        try:
            for key, value in data.items():
                if key == "ElementId":
                    continue

                try:
                    param = elem.LookupParameter(key)
                    if param and not param.IsReadOnly:
                        self._set_parameter_value(param, value)
                        updated = True
                except Exception as e:
                    print(f"Error updating parameter {key}: {e}")
                    continue

        except Exception as e:
            print(f"Error updating element parameters: {e}")

        return updated

    def _set_parameter_value(self, param, value):
        """Set the value of a parameter"""
        try:
            if param.StorageType.ToString() == "String":
                param.Set(str(value))
            elif param.StorageType.ToString() == "Double":
                param.Set(float(value))
            elif param.StorageType.ToString() == "Integer":
                param.Set(int(value))
            else:
                param.SetValueString(str(value))
        except Exception as e:
            print(f"Error setting parameter value: {e}")
