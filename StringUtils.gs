function toSnakeCase(string) {
  return string.trim().charAt(0).toLowerCase() + string.trim().slice(1) // Lowercase the first character
    .replace(/\W+/g, " ") // Remove all excess white space and replace & , . etc.
    .trim()
    .replace(/([a-z])([A-Z])([a-z])/g, "$1 $2$3") // Put a space at the position of a camelCase -> camel Case
    .split(/\B(?=[A-Z]{8,})/) // Now split the multi-uppercases customerID -> customer,ID
    .join(' ') // And join back with spaces.
    .split(' ') // Split all the spaces again, this time we're fully converted
    .join('_') // And finally snake_case things up
    .toLowerCase() // With a nice lower case
}

function testToSnakeCase() {
  Logger.log(toSnakeCase("ProteinaseK concentration"));
}
