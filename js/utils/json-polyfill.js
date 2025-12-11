/**
 * JSON Polyfill for IE/HTA compatibility
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
        parse: function(str) {
            return eval('(' + str + ')');
        }
    };
}
