using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ganss.Excel
{
    /// <summary>
    /// Enforce that result from formula is desired. Default true.
    /// </summary>
    /// <seealso cref="System.Attribute" />
    [AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
    public sealed class FormulaResultAttribute : Attribute
    {
        readonly bool asFormula = true;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormulaResultAttribute"/> class.
        /// </summary>
        /// <param asFormula="asFormula">Should formula output to destination</param>
        public FormulaResultAttribute()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormulaResultAttribute"/> class explicitly setting AsFormula value.
        /// </summary>
        /// <param asFormula="asFormula">Should formula output to destination</param>
        public FormulaResultAttribute(bool asFormula)
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
