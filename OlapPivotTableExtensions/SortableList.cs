using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OlapPivotTableExtensions
{

    public class SortableList<T> : BindingList<T>, IRaiseItemChangedEvents
    {

        //properties and methods of your business object list class

        //these are used to implement the sorting for the BindingList
        private bool m_Sorted = false;
        private ListSortDirection m_SortDirection = ListSortDirection.Ascending;
        private PropertyDescriptor m_SortProperty = null;

        //Sorting Code
        //To perform the sort of our BindingList<T> class, we have to provide the overrides
        //of all the sort-related methods and properties from the base class.

        /// <summary>
        /// Override the value and set it to true so sorting will work on the list.
        /// </summary>
        protected override bool SupportsSearchingCore
        {
            get
            {
                return true;
            }
        }

        /// <summary>
        /// Override the value and set it to true so sorting will work on the list.
        /// </summary>
        protected override bool SupportsSortingCore
        {
            get
            {
                return true;
            }
        }

        /// <summary>
        /// Return the value retained locally. Tells if the list is sorted or not.
        /// </summary>
        protected override bool IsSortedCore
        {
            get
            {
                return m_Sorted;
            }
        }

        /// <summary>
        /// Return the value retained locally. Tells which direction the list is sorted.
        /// </summary>
        protected override ListSortDirection SortDirectionCore
        {
            get
            {
                return m_SortDirection;
            }
        }

        /// <summary>
        /// Return the value retained locally. Tells which property the list is sorted on.
        /// </summary>
        protected override PropertyDescriptor SortPropertyCore
        {
            get
            {
                return m_SortProperty;
            }
        }

        /// <summary>
        /// Sets the properties when called by the base class in response to the ApplySort call.
        /// Delegates to a helper method (ApplySortInternal) to do most of the work of the sorting.
        /// </summary>
        protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
        {
            m_SortDirection = direction;
            m_SortProperty = prop;
            SortComparer<T> comparer = new SortComparer<T>(prop, direction);
            ApplySortInternal(comparer);
        }

        /// <summary>
        /// Helper class to do the actual sorting work.
        /// </summary>
        private void ApplySortInternal(SortComparer<T> comparer)
        {
            //this causes the items in the collection maintained by the base class to be sorted
            //  according to the criteria provided to the BOSortComparer class.
            List<T> listRef = this.Items as List<T>;
            if (listRef == null)
                return;

            //let List<T> do the actual sorting based on your comparer
            listRef.Sort(comparer);
            m_Sorted = true;
            //fire an event through a call to the base class OnListChanged method indicating
            //  that the list has been changed.
            //Use 'reset' because it's likely that most members have been moved around.
            OnListChanged(new
            ListChangedEventArgs(ListChangedType.Reset, -1));
        }

        /// <summary>
        /// Generic SortComparer classes. Used to sort a list of objects by any property.
        /// </summary>
        /// <typeparam name="U"></typeparam>
        class SortComparer<U> : IComparer<U>
        {
            private PropertyDescriptor m_PropDesc = null;
            private ListSortDirection m_Direction = ListSortDirection.Ascending;

            public SortComparer(PropertyDescriptor propDesc, ListSortDirection direction)
            {
                m_PropDesc = propDesc;
                m_Direction = direction;
            }

            int IComparer<U>.Compare(U x, U y)
            {
                object xValue = m_PropDesc.GetValue(x);
                object yValue = m_PropDesc.GetValue(y);
                return CompareValues(xValue, yValue, m_Direction);
            }

            private int CompareValues(object xValue, object yValue, ListSortDirection direction)
            {
                int retValue = 0;
                try
                {
                    if (xValue is IComparable) //can ask the x value
                    {
                        retValue = ((IComparable)xValue).CompareTo(yValue);
                    }
                    else if (yValue is IComparable) //can ask the y value
                    {
                        retValue = ((IComparable)yValue).CompareTo(xValue);
                    }
                    //not comparable, compare string representations
                    else if (xValue != null && yValue != null && !xValue.Equals(yValue))
                    {
                        retValue = xValue.ToString().CompareTo(yValue.ToString());
                    }
                }
                catch { }
                
                if (direction == ListSortDirection.Ascending)
                    return retValue;
                else
                    return retValue * -1;
            }

        }
    }
}
