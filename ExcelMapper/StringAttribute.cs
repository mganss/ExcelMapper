using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ganss.Excel
{
    /// <summary>
    /// Attribute that specifies how string type mapping to be handled.
    /// </summary>
    /// <seealso cref="System.Attribute" />
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public sealed class StringAttribute : Attribute
    {
        readonly bool asFormula;

        /// <summary>
        /// Initializes a new instance of the <see cref="StringAttribute"/> class.
        /// </summary>
        /// <param asFormula="asFormula">Should formula output to destination</param>
        public StringAttribute(bool asFormula)
        {
            this.asFormula = asFormula;
        }

        /// <summary>
        /// Should formula output to destination property
        /// </summary>
        /// <value>
        /// Boolean if destination should be the formula in the cell.
        /// </value>
        public bool AsFormula
        {
            get { return asFormula; }
        }
    }
}
