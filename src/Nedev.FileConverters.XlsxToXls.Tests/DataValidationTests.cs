using System;
using System.Collections.Generic;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class DataValidationTests
    {
        [Fact]
        public void DataValidationData_DefaultValues()
        {
            var validation = new DataValidationData();

            Assert.NotNull(validation.Id);
            Assert.NotNull(validation.Ranges);
            Assert.Empty(validation.Ranges);
            Assert.Equal(ValidationType.Any, validation.Type);
            Assert.Equal(ValidationOperator.Between, validation.Operator);
            Assert.Null(validation.Formula1);
            Assert.Null(validation.Formula2);
            Assert.True(validation.AllowBlank);
            Assert.False(validation.SuppressDropDown);
            Assert.True(validation.ShowInputMessage);
            Assert.True(validation.ShowErrorMessage);
            Assert.Null(validation.InputTitle);
            Assert.Null(validation.InputMessage);
            Assert.Null(validation.ErrorTitle);
            Assert.Null(validation.ErrorMessage);
            Assert.Equal(ErrorAlertType.Stop, validation.ErrorAlertType);
            Assert.Null(validation.ListValues);
            Assert.False(validation.AllowMultiSelect);
            Assert.False(validation.ShowErrorOnBlank);
        }

        [Fact]
        public void DataValidationData_WholeNumberValidation()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.Whole,
                Operator = ValidationOperator.GreaterThan,
                Formula1 = "0",
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                }
            };

            Assert.Equal(ValidationType.Whole, validation.Type);
            Assert.Equal(ValidationOperator.GreaterThan, validation.Operator);
            Assert.Equal("0", validation.Formula1);
            Assert.Single(validation.Ranges);
        }

        [Fact]
        public void DataValidationData_DecimalValidation()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.Decimal,
                Operator = ValidationOperator.Between,
                Formula1 = "0.0",
                Formula2 = "100.0",
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                }
            };

            Assert.Equal(ValidationType.Decimal, validation.Type);
            Assert.Equal(ValidationOperator.Between, validation.Operator);
            Assert.Equal("0.0", validation.Formula1);
            Assert.Equal("100.0", validation.Formula2);
        }

        [Fact]
        public void DataValidationData_ListValidation()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.List,
                Operator = ValidationOperator.None,
                ListValues = new List<string> { "Option1", "Option2", "Option3" },
                Formula1 = "Option1,Option2,Option3",
                SuppressDropDown = false,
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                }
            };

            Assert.Equal(ValidationType.List, validation.Type);
            Assert.Equal(3, validation.ListValues.Count);
            Assert.Equal("Option2", validation.ListValues[1]);
            Assert.False(validation.SuppressDropDown);
        }

        [Fact]
        public void DataValidationData_DateValidation()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.Date,
                Operator = ValidationOperator.GreaterThanOrEqual,
                Formula1 = "44561", // Excel date serial number
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                }
            };

            Assert.Equal(ValidationType.Date, validation.Type);
            Assert.Equal(ValidationOperator.GreaterThanOrEqual, validation.Operator);
            Assert.Equal("44561", validation.Formula1);
        }

        [Fact]
        public void DataValidationData_TimeValidation()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.Time,
                Operator = ValidationOperator.Between,
                Formula1 = "0.25", // 6:00 AM (quarter of a day)
                Formula2 = "0.75", // 6:00 PM (three quarters of a day)
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                }
            };

            Assert.Equal(ValidationType.Time, validation.Type);
            Assert.Equal(ValidationOperator.Between, validation.Operator);
            Assert.Equal("0.25", validation.Formula1);
            Assert.Equal("0.75", validation.Formula2);
        }

        [Fact]
        public void DataValidationData_TextLengthValidation()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.TextLength,
                Operator = ValidationOperator.LessThanOrEqual,
                Formula1 = "50",
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                }
            };

            Assert.Equal(ValidationType.TextLength, validation.Type);
            Assert.Equal(ValidationOperator.LessThanOrEqual, validation.Operator);
            Assert.Equal("50", validation.Formula1);
        }

        [Fact]
        public void DataValidationData_CustomValidation()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.Custom,
                Operator = ValidationOperator.None,
                Formula1 = "=A1>B1",
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                }
            };

            Assert.Equal(ValidationType.Custom, validation.Type);
            Assert.Equal(ValidationOperator.None, validation.Operator);
            Assert.Equal("=A1>B1", validation.Formula1);
        }

        [Fact]
        public void DataValidationData_WithInputMessage()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.Whole,
                Operator = ValidationOperator.Between,
                Formula1 = "1",
                Formula2 = "100",
                InputTitle = "Enter a Number",
                InputMessage = "Please enter a number between 1 and 100.",
                ShowInputMessage = true,
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                }
            };

            Assert.Equal("Enter a Number", validation.InputTitle);
            Assert.Equal("Please enter a number between 1 and 100.", validation.InputMessage);
            Assert.True(validation.ShowInputMessage);
        }

        [Fact]
        public void DataValidationData_WithErrorMessage()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.Whole,
                Operator = ValidationOperator.Between,
                Formula1 = "1",
                Formula2 = "100",
                ErrorTitle = "Invalid Input",
                ErrorMessage = "The value must be between 1 and 100.",
                ErrorAlertType = ErrorAlertType.Stop,
                ShowErrorMessage = true,
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                }
            };

            Assert.Equal("Invalid Input", validation.ErrorTitle);
            Assert.Equal("The value must be between 1 and 100.", validation.ErrorMessage);
            Assert.Equal(ErrorAlertType.Stop, validation.ErrorAlertType);
            Assert.True(validation.ShowErrorMessage);
        }

        [Fact]
        public void DataValidationData_WithWarningAlert()
        {
            var validation = new DataValidationData
            {
                ErrorAlertType = ErrorAlertType.Warning,
                ErrorTitle = "Warning",
                ErrorMessage = "Value is outside the recommended range."
            };

            Assert.Equal(ErrorAlertType.Warning, validation.ErrorAlertType);
        }

        [Fact]
        public void DataValidationData_WithInformationAlert()
        {
            var validation = new DataValidationData
            {
                ErrorAlertType = ErrorAlertType.Information,
                ErrorTitle = "Information",
                ErrorMessage = "Please note the value range."
            };

            Assert.Equal(ErrorAlertType.Information, validation.ErrorAlertType);
        }

        [Fact]
        public void DataValidationData_MultipleRanges()
        {
            var validation = new DataValidationData
            {
                Type = ValidationType.List,
                ListValues = new List<string> { "Yes", "No" },
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 10, LastCol = 0 },
                    new() { FirstRow = 0, FirstCol = 2, LastRow = 10, LastCol = 2 },
                    new() { FirstRow = 0, FirstCol = 4, LastRow = 10, LastCol = 4 }
                }
            };

            Assert.Equal(3, validation.Ranges.Count);
        }

        // Validation type tests
        [Theory]
        [InlineData(ValidationType.Any, 0)]
        [InlineData(ValidationType.Whole, 1)]
        [InlineData(ValidationType.Decimal, 2)]
        [InlineData(ValidationType.List, 3)]
        [InlineData(ValidationType.Date, 4)]
        [InlineData(ValidationType.Time, 5)]
        [InlineData(ValidationType.TextLength, 6)]
        [InlineData(ValidationType.Custom, 7)]
        public void ValidationType_HasCorrectValues(ValidationType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        // Validation operator tests
        [Theory]
        [InlineData(ValidationOperator.None, 0)]
        [InlineData(ValidationOperator.Between, 1)]
        [InlineData(ValidationOperator.NotBetween, 2)]
        [InlineData(ValidationOperator.Equal, 3)]
        [InlineData(ValidationOperator.NotEqual, 4)]
        [InlineData(ValidationOperator.GreaterThan, 5)]
        [InlineData(ValidationOperator.LessThan, 6)]
        [InlineData(ValidationOperator.GreaterThanOrEqual, 7)]
        [InlineData(ValidationOperator.LessThanOrEqual, 8)]
        public void ValidationOperator_HasCorrectValues(ValidationOperator op, byte expected)
        {
            Assert.Equal(expected, (byte)op);
        }

        // Error alert type tests
        [Theory]
        [InlineData(ErrorAlertType.Stop, 0)]
        [InlineData(ErrorAlertType.Warning, 1)]
        [InlineData(ErrorAlertType.Information, 2)]
        public void ErrorAlertType_HasCorrectValues(ErrorAlertType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        // DataValidationHelper tests
        [Fact]
        public void DataValidationHelper_CreateWholeNumberValidation()
        {
            var ranges = new List<CellRange>
            {
                new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
            };

            var validation = DataValidationHelper.CreateWholeNumberValidation(
                ranges, ValidationOperator.GreaterThan, 0, null);

            Assert.Equal(ValidationType.Whole, validation.Type);
            Assert.Equal(ValidationOperator.GreaterThan, validation.Operator);
            Assert.Equal("0", validation.Formula1);
            Assert.Null(validation.Formula2);
        }

        [Fact]
        public void DataValidationHelper_CreateDecimalValidation()
        {
            var ranges = new List<CellRange>
            {
                new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
            };

            var validation = DataValidationHelper.CreateDecimalValidation(
                ranges, ValidationOperator.Between, 0.0, 100.0);

            Assert.Equal(ValidationType.Decimal, validation.Type);
            Assert.Equal(ValidationOperator.Between, validation.Operator);
            Assert.Equal("0", validation.Formula1);
            Assert.Equal("100", validation.Formula2);
        }

        [Fact]
        public void DataValidationHelper_CreateListValidation()
        {
            var ranges = new List<CellRange>
            {
                new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
            };
            var values = new List<string> { "Red", "Green", "Blue" };

            var validation = DataValidationHelper.CreateListValidation(ranges, values, false, true);

            Assert.Equal(ValidationType.List, validation.Type);
            Assert.Equal(3, validation.ListValues.Count);
            Assert.False(validation.AllowBlank);
            Assert.True(validation.SuppressDropDown);
        }

        [Fact]
        public void DataValidationHelper_CreateTextLengthValidation()
        {
            var ranges = new List<CellRange>
            {
                new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
            };

            var validation = DataValidationHelper.CreateTextLengthValidation(
                ranges, ValidationOperator.LessThanOrEqual, null, 50);

            Assert.Equal(ValidationType.TextLength, validation.Type);
            Assert.Equal(ValidationOperator.LessThanOrEqual, validation.Operator);
            Assert.Equal("50", validation.Formula2);
        }

        [Fact]
        public void DataValidationHelper_CreateCustomValidation()
        {
            var ranges = new List<CellRange>
            {
                new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
            };

            var validation = DataValidationHelper.CreateCustomValidation(ranges, "=A1>B1");

            Assert.Equal(ValidationType.Custom, validation.Type);
            Assert.Equal("=A1>B1", validation.Formula1);
        }

        [Fact]
        public void DataValidationHelper_WithInputMessage()
        {
            var validation = new DataValidationData();
            validation.WithInputMessage("Title", "Message");

            Assert.Equal("Title", validation.InputTitle);
            Assert.Equal("Message", validation.InputMessage);
            Assert.True(validation.ShowInputMessage);
        }

        [Fact]
        public void DataValidationHelper_WithErrorMessage()
        {
            var validation = new DataValidationData();
            validation.WithErrorMessage("Error", "Invalid value", ErrorAlertType.Warning);

            Assert.Equal("Error", validation.ErrorTitle);
            Assert.Equal("Invalid value", validation.ErrorMessage);
            Assert.Equal(ErrorAlertType.Warning, validation.ErrorAlertType);
            Assert.True(validation.ShowErrorMessage);
        }

        [Fact]
        public void DataValidationHelper_ToR1C1Reference()
        {
            Assert.Equal("R1C1", DataValidationHelper.ToR1C1Reference(0, 0));
            Assert.Equal("R10C5", DataValidationHelper.ToR1C1Reference(9, 4));
            Assert.Equal("R100C26", DataValidationHelper.ToR1C1Reference(99, 25));
        }

        [Fact]
        public void DataValidationHelper_ToR1C1Range()
        {
            var range = new CellRange
            {
                FirstRow = 0,
                FirstCol = 0,
                LastRow = 9,
                LastCol = 4
            };

            Assert.Equal("R1C1:R10C5", DataValidationHelper.ToR1C1Range(range));
        }

        // DataValidationWriter tests
        [Fact]
        public void DataValidationWriter_CreatePooled()
        {
            var writer = DataValidationWriter.CreatePooled(out var buffer, 8192);
            try
            {
                Assert.True(buffer.Length >= 8192);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void DataValidationWriter_WritesEmptyList()
        {
            var writer = DataValidationWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var validations = new List<DataValidationData>();
                var bytesWritten = writer.WriteDataValidations(validations, 0);
                Assert.Equal(0, bytesWritten);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void DataValidationWriter_WritesWholeNumberValidation()
        {
            var writer = DataValidationWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var validations = new List<DataValidationData>
                {
                    new()
                    {
                        Type = ValidationType.Whole,
                        Operator = ValidationOperator.GreaterThan,
                        Formula1 = "0",
                        Ranges = new List<CellRange>
                        {
                            new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                        }
                    }
                };

                var bytesWritten = writer.WriteDataValidations(validations, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void DataValidationWriter_WritesListValidation()
        {
            var writer = DataValidationWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var validations = new List<DataValidationData>
                {
                    new()
                    {
                        Type = ValidationType.List,
                        ListValues = new List<string> { "Yes", "No", "Maybe" },
                        Formula1 = "Yes,No,Maybe",
                        Ranges = new List<CellRange>
                        {
                            new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                        }
                    }
                };

                var bytesWritten = writer.WriteDataValidations(validations, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void DataValidationWriter_WritesValidationWithMessages()
        {
            var writer = DataValidationWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var validations = new List<DataValidationData>
                {
                    new()
                    {
                        Type = ValidationType.Decimal,
                        Operator = ValidationOperator.Between,
                        Formula1 = "0",
                        Formula2 = "100",
                        InputTitle = "Enter Value",
                        InputMessage = "Enter a value between 0 and 100",
                        ErrorTitle = "Invalid",
                        ErrorMessage = "Value must be between 0 and 100",
                        ErrorAlertType = ErrorAlertType.Stop,
                        Ranges = new List<CellRange>
                        {
                            new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                        }
                    }
                };

                var bytesWritten = writer.WriteDataValidations(validations, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void DataValidationWriter_WritesMultipleValidations()
        {
            var writer = DataValidationWriter.CreatePooled(out var buffer, 16384);
            try
            {
                var validations = new List<DataValidationData>
                {
                    new()
                    {
                        Type = ValidationType.Whole,
                        Operator = ValidationOperator.GreaterThan,
                        Formula1 = "0",
                        Ranges = new List<CellRange>
                        {
                            new() { FirstRow = 0, FirstCol = 0, LastRow = 50, LastCol = 0 }
                        }
                    },
                    new()
                    {
                        Type = ValidationType.List,
                        ListValues = new List<string> { "A", "B", "C" },
                        Formula1 = "A,B,C",
                        Ranges = new List<CellRange>
                        {
                            new() { FirstRow = 0, FirstCol = 1, LastRow = 50, LastCol = 1 }
                        }
                    },
                    new()
                    {
                        Type = ValidationType.Date,
                        Operator = ValidationOperator.GreaterThanOrEqual,
                        Formula1 = "44561",
                        Ranges = new List<CellRange>
                        {
                            new() { FirstRow = 0, FirstCol = 2, LastRow = 50, LastCol = 2 }
                        }
                    }
                };

                var bytesWritten = writer.WriteDataValidations(validations, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void DataValidationWriter_WritesMultipleRanges()
        {
            var writer = DataValidationWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var validations = new List<DataValidationData>
                {
                    new()
                    {
                        Type = ValidationType.List,
                        ListValues = new List<string> { "Yes", "No" },
                        Formula1 = "Yes,No",
                        Ranges = new List<CellRange>
                        {
                            new() { FirstRow = 0, FirstCol = 0, LastRow = 10, LastCol = 0 },
                            new() { FirstRow = 0, FirstCol = 2, LastRow = 10, LastCol = 2 }
                        }
                    }
                };

                var bytesWritten = writer.WriteDataValidations(validations, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void DataValidationWriter_WritesCustomValidation()
        {
            var writer = DataValidationWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var validations = new List<DataValidationData>
                {
                    new()
                    {
                        Type = ValidationType.Custom,
                        Operator = ValidationOperator.None,
                        Formula1 = "=A1>B1",
                        Ranges = new List<CellRange>
                        {
                            new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                        }
                    }
                };

                var bytesWritten = writer.WriteDataValidations(validations, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }
    }
}
