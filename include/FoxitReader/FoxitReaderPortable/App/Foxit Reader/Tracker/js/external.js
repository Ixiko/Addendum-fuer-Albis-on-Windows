;(function (global) {
    var external = window.external;

    global.externalDispatchFun = function () {
        try {
            return DispatchFun(arguments);
        } catch (e) {
            try {
                return dispatchFun(arguments);
            } catch (e) {
                throw e;
            }
        }
    };

    function dispatchFun (args) {
        switch (args.length) {
            case 2:
                return external.dispatchFun(args[0], args[1]);
            case 3:
                return external.dispatchFun(args[0], args[1], args[2]);
            default :
                throw new Error('argument number error');
        }
    }
    function DispatchFun (args) {
        switch (args.length) {
            case 2:
                return external.DispatchFun(args[0], args[1]);
            case 3:
                return external.DispatchFun(args[0], args[1], args[2]);
            default :
                throw new Error('argument number error');
        }
    }

    var _alert = alert;

    global.notify = function (msg) {
        var div = document.createElement('div');
        div.style = 'z-index: 100000; color: red; font-size: 50px; position: fixed; top: 0; left: 0;';
        var body = document.getElementsByTagName('body')[0];
        body.appendChild(div);
        try {
            div.innerHTML = '1';
            window.external.notify(msg);
            div.innerHTML = '2';
        } catch (e) {
            div.innerHTML = '3';
            _alert(msg);
            div.innerHTML = '4';
        }
    }

})(this);