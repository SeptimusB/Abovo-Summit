# Questions for the client

## Business Plan Start Date

Should the Business Plan Start Date:

1. be restricted to the first day of a month; or
2. support any calendar date?

The current workbook presents these as conflicting contracts. Excel validation on `Global Assumptions!C8` permits any date after 1 January 2024, but downstream monthly period tables and exact-match formulas require a month-start date. For example, entering 26 April 2025 produces `#N/A` values through Development BP Assumptions, Development Expenditure, Transactional DB, Cashflow and SOFP.

The client's answer determines the correct implementation:

- If month-start dates are required, update the authoritative XLSB validation and Summit editor validation consistently.
- If any calendar date must be supported, revise the authoritative XLSB period/date formulas and test Excel-Summit-Excel round-tripping.
