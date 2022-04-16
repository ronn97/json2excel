export default class Util {
    static createUUID = () => {
        return 'xxxxxxxxxxxx8xxxyxxxxxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
            var r = Math.random() * 16 | 0, v = (c === 'x') ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    };

    static addArrayItems = (items, its) => {
        for (let i = 0; i < its.length; ++i) {
            const item = its[i];
            items = Util.addArrayItem(items, item);
        }
        return items;
    };

    static addArrayItem = (items, item) => {
        for (let i = 0; i < items.length; ++i) {
            if (items[i] === item) {
                return items;
            }
        }
        items.push(item);
        return items;
    };

    static webGetImage = (url, callback) => {
        fetch(url, {
            cache: 'no-cache'
        })
            .then((res) => {
                return res.blob();
            }).then(blob => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    callback(e.target.result);
                };
                reader.readAsDataURL(blob);
            });
    };

    static handleMultiObjects = (objects, handleOneObject, params, callbackOfAllDone = undefined, callbackOfOneDone = undefined, index = undefined) => {
        if (index === undefined) {
            index = 0;
        }
        if (index >= objects.length) {
            if (callbackOfAllDone) {
                callbackOfAllDone(objects);
            }
            return;
        }
        const obj = objects[index];
        handleOneObject(obj, params, data => {
            if (callbackOfOneDone) {
                callbackOfOneDone(data, index);
            }
            Util.handleMultiObjects(objects, handleOneObject, params, callbackOfAllDone, callbackOfOneDone, index + 1);
        });
    };

}