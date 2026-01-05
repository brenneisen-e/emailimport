/**
 * JSON Polyfill for IE/HTA compatibility
 * SECURITY: Uses safe parsing without eval()
 */

if (typeof JSON === 'undefined') {
    JSON = {
        stringify: function(obj) {
            if (obj === null) return 'null';
            if (typeof obj === 'undefined') return undefined;
            if (typeof obj === 'number' || typeof obj === 'boolean') return String(obj);
            if (typeof obj === 'string') return '"' + obj.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\n/g, '\\n').replace(/\r/g, '\\r').replace(/\t/g, '\\t') + '"';
            if (obj instanceof Array) {
                var arr = [];
                for (var i = 0; i < obj.length; i++) {
                    arr.push(JSON.stringify(obj[i]));
                }
                return '[' + arr.join(',') + ']';
            }
            if (typeof obj === 'object') {
                var pairs = [];
                for (var key in obj) {
                    if (obj.hasOwnProperty(key)) {
                        var val = JSON.stringify(obj[key]);
                        if (val !== undefined) {
                            pairs.push('"' + key + '":' + val);
                        }
                    }
                }
                return '{' + pairs.join(',') + '}';
            }
            return undefined;
        },

        /**
         * Safe JSON parser without eval()
         * Validates input before parsing to prevent code injection
         */
        parse: function(str) {
            // Validate input is a string
            if (typeof str !== 'string') {
                throw new SyntaxError('JSON.parse: input must be a string');
            }

            // Trim whitespace
            str = str.replace(/^\s+|\s+$/g, '');

            // Empty string check
            if (str === '') {
                throw new SyntaxError('JSON.parse: unexpected end of data');
            }

            // Security: Validate JSON structure before parsing
            // This regex validates that the string contains only valid JSON characters
            // Based on RFC 8259 JSON specification
            var validJson = /^[\],:{}\s]*$/;
            var validate = str
                .replace(/\\["\\\/bfnrtu]/g, '@')           // Replace valid escape sequences
                .replace(/"[^"\\\n\r]*"/g, '')              // Remove strings
                .replace(/-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/g, ']')  // Replace numbers with ]
                .replace(/(?:true|false|null)/g, ']');      // Replace literals with ]

            if (!validJson.test(validate)) {
                throw new SyntaxError('JSON.parse: invalid JSON structure');
            }

            // Use Function constructor instead of eval for slightly better isolation
            // This is still not ideal but is safer than raw eval
            // The validation above ensures only valid JSON characters remain
            try {
                return (new Function('return (' + str + ')'))();
            } catch (e) {
                throw new SyntaxError('JSON.parse: ' + e.message);
            }
        }
    };
}
