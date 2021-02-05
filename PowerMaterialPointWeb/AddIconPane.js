'use strict';

(function () {

    const sets = {
        'material': new MaterialSet()
    };

    const currentSetName = 'material',
        currentSet = sets[currentSetName],
        optionsFactory = {},
        messages = {};

    function init() {
        const iconsDatabaseLoadingPromise = fetch(`${currentSetName}_icons.json`)
            .then(response => response.json())
            .then(iconDb => currentSet.db = iconDb);
        const officeReadyPromise = new Promise(resolve => {
            Office.onReady(resolve);
        });
        const domContentLoadedPromise = new Promise(resolve => {
            document.addEventListener('DOMContentLoaded', () => resolve());
        });

        Promise.all([officeReadyPromise, domContentLoadedPromise, iconsDatabaseLoadingPromise])
            .then(() => iconsDbLoaded());
    }

    function $id(id) {
        return document.getElementById(id);
    }

    function $(selector) {
        return document.querySelector(selector);
    }

    function create(element) {
        return document.createElement(element);
    }

    function $all(selector) {
        return document.querySelectorAll(selector);
    }

    function renderMessages() {
        const msgs = $id('messages');

        if (msgs.innerHTML)
            msgs.innerHTML = '';

        if (currentSet.messages && currentSet.messages.length) {
            currentSet.messages.forEach(msg=> {
                const msgDiv = create('div');
                msgDiv.innerText = msg;
                msgs.appendChild(msgDiv);
            });
        }            
    }

    function blobToBase64Promise(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result.substr(reader.result.indexOf(',') + 1));
            reader.onerror = e => reject(e);
            reader.readAsDataURL(blob);
        });
    }

    function iconsDbLoaded() {
        const iconsSection = $id('icons');

        currentSet.setup(iconsSection);

        iconsSection.classList.add(currentSetName);

        setupOptions();

        $id('loading-icons').remove();

        optionsChanged();

        Object.keys(currentSet.db).forEach(categoryName => {
            const category = create('fieldset'),
                legend = create('legend');


            legend.innerText = categoryName;
            category.appendChild(legend);

            category.classList.add('icons-category');
            iconsSection.appendChild(category);

            const icons = currentSet.db[categoryName];
            let buttons;

            if (Array.isArray(icons))
                buttons = icons.map((name, index) => createIconButton(categoryName, index, name));
            else
                buttons = Object.keys(icons).map(name => createIconButton(categoryName, name, icons[name]));

            buttons.forEach(category.appendChild, category);

            category.addEventListener('click', insertIconButtonClick);
        });
    }

    function setupOptions() {
        const optionsSection = $id('options'),
            hasOptions = !!Object.keys(currentSet.options).length;

        if (hasOptions) {
            Object.keys(currentSet.options)
                .forEach(option => {
                    if (option == 'color')
                        optionsFactory.color = setupColorPicker(optionsSection);
                    else if (Array.isArray(currentSet.options[option]))
                        optionsFactory[option] = setupDropdownOptions(optionsSection, option, currentSet.options[option]);
                })
        }
    }

    function setupDropdownOptions(optionsSection, name, values) {
        const select = create('select');
        optionsSection.appendChild(select);

        select.title = name;
        values.forEach(value => {
            const opt = create('option');
            opt.innerText = value;
            opt.value = value;
            select.appendChild(opt);
        });

        select.addEventListener('input', () => optionsChanged());

        return () => select.value;
    }

    function setupColorPicker(optionsSection) {
        const colorPickerDiv = create('div');
        const initColor = 'black';
        colorPickerDiv.id = 'color-picker';
        colorPickerDiv.style.background = initColor;
        optionsSection.appendChild(colorPickerDiv);
        const picker = new Picker({
            parent: colorPickerDiv,
            color: initColor,
            alpha: false,
            onDone: color => {
                colorPickerDiv.style.background = color.rgbString
                optionsChanged();
            }
        });

        return () => colorPickerDiv.style.background;
    }

    function getCurrentOptions() {
        const options = {};

        Object.keys(optionsFactory).forEach(optionName => options[optionName] = optionsFactory[optionName]());

        return options;
    }

    function createIconButton(iconCategory, iconName, iconDefinition) {
        const iconButton = create('button');
        iconButton.dataset.category = iconCategory;
        iconButton.dataset.name = iconName;
        iconButton.title = iconName;

        currentSet.setupButton(iconButton, iconName, iconDefinition, iconCategory);

        return iconButton;
    }

    function optionsChanged() {
        currentSet.onOptionsChange(getCurrentOptions(), $id('icons'));
        renderMessages();
    }

    function insertIconButtonClick(event) {
        if (event.target instanceof HTMLButtonElement) {
            const button = event.target;
            currentSet.fetchIcon(button.dataset.name, currentSet.db[button.dataset.category][button.dataset.name], getCurrentOptions())
                .then(image => {
                    Office.context.document.setSelectedDataAsync(image, {
                        coercionType: Office.CoercionType[currentSet.imageType],
                        imageLeft: 50,
                        imageTop: 50,
                        imageWidth: 400
                    });
                });
        }
    }

    init();


    function MaterialSet() {

        const set = this,
            themes = {
                sharp: 'Sharp',
                filled: 'Filled',
                outlined: 'Outlined',
                rounded: 'Rounded',
                twoTone: 'Two-Tone',
            };
        
        let messages = [];

        Object.defineProperty(set, 'messages', {
            get: () => messages
        });

        set.imageType = 'XmlSvg';

        set.setupButton = (button, name, version, category) => button.innerText = name;

        set.fetchIcon = function (name, version, options) {
            let themeFolder;

            switch (options.theme) {
                case themes.sharp:
                    themeFolder = 'materialiconssharp';
                    break;

                case themes.filled:
                    themeFolder = 'materialicons';
                    break;

                case themes.outlined:
                    themeFolder = 'materialiconsoutlined';
                    break;

                case themes.rounded:
                    themeFolder = 'materialiconsround';
                    break;

                case themes.twoTone:
                    themeFolder = 'materialiconstwotone';
                    break;                    
            }

            return fetch(`https://fonts.gstatic.com/s/i/${themeFolder}/${name}/v${version}/24px.svg`)
                .then(response => response.text())
                .then(svgString => {
                    const svg = new DOMParser().parseFromString(svgString, "image/svg+xml");
                    svg.rootElement.style.fill = options.color;
                    return svg.rootElement.outerHTML;
                });
        }

        set.options = {
            color: true,
            theme: Object.values(themes)
        };

        set.onOptionsChange = function (newOptions, iconsSection) {
            updateTheme(newOptions.theme, iconsSection);
            updateColor(newOptions.color, iconsSection);

            if (newOptions.theme === themes.twoTone && newOptions.color !== 'black' && newOptions.color !== 'rgb(0, 0, 0)') {
                messages = ["Os ícones serão inseridos com a cor correta apesar de alguns estarem sempre pretos abaixo"];
            }
            else
                messages = [];

        };

        set.setup = function (iconsSection) {
            const fontStyle = create('link');
            fontStyle.rel = 'stylesheet';
            fontStyle.href = 'https://fonts.googleapis.com/css?family=Material+Icons|Material+Icons+Outlined|Material+Icons+Sharp|Material+Icons+Round|Material+Icons+Two+Tone'
            document.head.appendChild(fontStyle);
        }


        function updateColor(color, iconsSection) {
            iconsSection.style.color = color;
        }

        function updateTheme(theme, iconsSection) {
            let className;

            iconsSection.classList.remove('mat-fill', 'mat-outlined', 'mat-round', 'mat-two-tone', 'mat-sharp');

            switch (theme) {
                case themes.filled:
                    className = 'mat-fill';
                    break;

                case themes.outlined:
                    className = 'mat-outlined';
                    break;

                case themes.rounded:
                    className = 'mat-round';
                    break;

                case themes.twoTone:
                    className = 'mat-two-tone';
                    break;

                default:
                    className = 'mat-sharp';
                    break;
            }

            iconsSection.classList.add(className);
        }
    }
})();