using System.Collections.Generic;
using BusinessCardMaker.Core.Models;
using Xunit;

namespace BusinessCardMaker.Tests.Models;

public class ImportResultTests
{
    [Fact]
    public void CreateSuccess_PopulatesEmployeesAndWarnings()
    {
        var employees = new List<Employee>
        {
            new() { Name = "Alice", Email = "alice@example.com" }
        };
        var warnings = new List<string> { "Sample warning" };

        var result = ImportResult.CreateSuccess(employees, warnings);

        Assert.True(result.Success);
        Assert.Equal(1, result.TotalCount);
        Assert.Same(employees, result.Employees);
        Assert.Same(warnings, result.Warnings);
        Assert.Null(result.ErrorMessage);
    }

    [Fact]
    public void CreateFailure_SetsErrorMessageAndClearsEmployees()
    {
        var result = ImportResult.CreateFailure("failure");

        Assert.False(result.Success);
        Assert.Equal("failure", result.ErrorMessage);
        Assert.Empty(result.Employees);
        Assert.Empty(result.Warnings);
    }
}
