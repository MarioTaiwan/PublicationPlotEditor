function PublicationPlotEditor
% PUBLICATIONPLOTEDITOR  Reproducible Excel-to-CorelDRAW plot editor.
% Version 0.7, intended for MATLAB R2021a+ on Windows.
% Start by running: PublicationPlotEditor
%
% Main concepts:
%   * Paste an Excel chart =SERIES(...) formula, using a selected workbook.
%   * Search writable MATLAB graphics properties; add them to 'My Editor'.
%   * Import RGB/sRGB CorelDRAW XML palette colors (copied, not linked).
%   * Save figure recipe JSON + cached source data MAT file.
%   * Save/load a master template; export dimensioned EMF/PDF vector plots.
%
% A few graphics properties are MATLAB-release dependent. The Advanced
% property browser only allows properties MATLAB reports as settable.
%
% This prototype focuses on standard 2D line/marker plots. Native graphics
% properties accepting handles/callbacks are shown but not directly editable.

    S = struct();
    S.appVersion = '0.7';
    S.masterName = 'Standard Scientific';
    S.masterNotes = '';
    S.workbook = '';
    S.datasets = struct('name',{},'formula',{},'workbook',{}, ...
        'xSheet',{},'xRange',{},'ySheet',{},'yRange',{}, ...
        'order',{},'x',{},'y',{},'style',{},'sourceColorName',{});
    S.activeSeries = 1;
    % widthMM and heightMM are the INNER PLOT BOX, not export canvas.
    % The complete EMF/PDF bounds include the four physical margins.
    S.panel = struct('preset','normal','dimensionMode','plotArea', ...
        'widthMM',85,'heightMM',65, ...
        'leftMM',15,'rightMM',5,'bottomMM',13,'topMM',5);
    % Separate X/Y tick and axis-label font corrections for CorelDRAW 19.
    % Stored in master templates; never applied to PDF or live preview.
    S.export = defaultExportSettings();
    S.custom = defaultCustomSettings();
    % All axis limits remain in S.overrides.Axes so there is one source
    % of truth for the native property inspector and the quick controls.
    S.ticks=struct('X',defaultTickSettings(),'Y',defaultTickSettings());
    S.palettePage=1;
    S.paletteSelected=1;
    S.palettePendingKey='';
    S.overrides = struct('Axes',struct(),'XRuler',struct(), ...
        'YRuler',struct(),'XLabel',struct('String','X'), ...
        'YLabel',struct('String','Y'),'Legend',struct());
    S.styleDefaults = defaultStyle();
    S.legendOn = false;
    S.palette = struct('path','','name','Built-in', ...
        'colors',defaultPalette());
    % My Editor starts with common dataset styling. Axis text/limits/ticks
    % and panel dimensions are controlled in their respective tabs.
    S.workspace = {'Dataset|Marker','Dataset|MarkerSize', ...
        'Dataset|LineStyle','Dataset|LineWidth', ...
        'Dataset|Color','Dataset|MarkerFaceColor', ...
        'Dataset|MarkerEdgeColor','Dataset|Visible'};
    S.recipePath = '';
    S.catalog = struct('key',{},'category',{},'property',{},'editable',{});
    S.displayCatalog = {};
    S.editorSelection = [];
    S.plotHandles = {};
    S.lastStatus = '';

    % Standard MATLAB figures/uicontrols are used for broad compatibility.
    ui = struct();
    ui.fig = figure('Name','Publication Plot Editor v0.7', ...
        'NumberTitle','off','MenuBar','none','ToolBar','none', ...
        'Units','pixels','Position',[50 45 1280 760], ...
        'Color',[0.95 0.96 0.97], 'Resize','on');

    ui.title = uicontrol(ui.fig,'Style','text','String','Plot Editor', ...
        'Units','normalized','Position',[.018 .947 .15 .039], ...
        'BackgroundColor',get(ui.fig,'Color'),'FontSize',14, ...
        'FontWeight','bold','HorizontalAlignment','left');
    % Master design controls are always visible in the main toolbar.
    makeText(ui.fig,'Master:',[.170 .952 .061 .031]);
    ui.masterName = uicontrol(ui.fig,'Style','edit', ...
        'String',S.masterName,'Units','normalized', ...
        'Position',[.232 .951 .157 .034], ...
        'HorizontalAlignment','left', 'Callback',@masterNameEdited, ...
        'TooltipString','Name of the currently active master design');
    makeButton(ui.fig,'Load master', [.395 .951 .104 .034],@loadTemplate);
    makeButton(ui.fig,'Save master', [.505 .951 .103 .034],@saveTemplate);
    makeButton(ui.fig,'Open figure', [.615 .951 .104 .034],@loadRecipe);
    makeButton(ui.fig,'Save figure', [.725 .951 .105 .034],@saveRecipe);
    makeButton(ui.fig,'Export', [.838 .951 .143 .034],@exportFigure);
    ui.previewPanel = uipanel('Parent',ui.fig,'Title','Live preview', ...
        'Units','normalized','Position',[.018 .295 .565 .636]);
    ui.ax = axes('Parent',ui.previewPanel,'Units','normalized', ...
        'Position',[.16 .19 .77 .72],'Box','off');
    ui.sourcePanel = uipanel('Parent',ui.fig,'Title','Excel series input', ...
        'Units','normalized','Position',[.018 .078 .565 .206]);
    makeText(ui.sourcePanel,'Workbook:',[.018 .79 .11 .14]);
    ui.workbook = uicontrol(ui.sourcePanel,'Style','edit','String','', ...
        'Units','normalized','Position',[.13 .77 .70 .18], ...
        'HorizontalAlignment','left','Callback',@workbookEdited);
    makeButton(ui.sourcePanel,'Browse...', [.84 .77 .145 .18], @browseWorkbook);
    makeText(ui.sourcePanel,'Paste one or more =SERIES(...) formulas:', ...
        [.018 .52 .66 .16]);
    ui.formula = uicontrol(ui.sourcePanel,'Style','edit','Max',12,'Min',0, ...
        'Units','normalized','Position',[.018 .13 .76 .42], ...
        'HorizontalAlignment','left','String','');
    makeButton(ui.sourcePanel,'Add series', [.797 .31 .185 .22], @addSeries);
    makeButton(ui.sourcePanel,'Refresh Excel', [.797 .07 .185 .20], @refreshExcel);

    ui.tabs = uitabgroup('Parent',ui.fig,'Units','normalized', ...
        'Position',[.595 .078 .389 .853]);
    ui.tabEditor = uitab(ui.tabs,'Title','My Editor');
    ui.tabProps = uitab(ui.tabs,'Title','Properties');
    ui.tabSeries = uitab(ui.tabs,'Title','Datasets');
    ui.tabPalette = uitab(ui.tabs,'Title','Palette');
    ui.tabPanel = uitab(ui.tabs,'Title','Panel');
    ui.tabTicks = uitab(ui.tabs,'Title','Axes & ticks');

    % A selected property is edited through either a dropdown (enumerations),
    % a palette selector (colors), or a text/numeric field.
    makeText(ui.tabEditor,'Selected series:',[.035 .933 .25 .035]);
    ui.editorSeries=uicontrol(ui.tabEditor,'Style','popupmenu', ...
        'String',{'(No series)'},'Units','normalized', ...
        'Position',[.30 .928 .65 .043],'Callback',@editorSeriesChanged);
    ui.editor=uitable(ui.tabEditor,'Units','normalized', ...
        'Position',[.035 .333 .93 .574], 'Data',cell(0,2), ...
        'ColumnName',{'Property','Current value'}, ...
        'ColumnWidth',{230,190},'RowName',[], ...
        'ColumnEditable',[false false], ...
        'CellSelectionCallback',@editorCellSelected);
    ui.editorHelp=makeText(ui.tabEditor,'Click a property to edit.', ...
        [.035 .273 .93 .035]);
    ui.valueEdit=uicontrol(ui.tabEditor,'Style','edit','String','', ...
        'Units','normalized','Position',[.035 .214 .61 .052], ...
        'HorizontalAlignment','left','Visible','off');
    ui.valueChoice=uicontrol(ui.tabEditor,'Style','popupmenu', ...
        'String',{'(select)'},'Units','normalized', ...
        'Position',[.035 .214 .61 .052],'Visible','off');
    ui.quickPalette=uicontrol(ui.tabEditor,'Style','popupmenu', ...
        'String',{'(no colors)'},'Units','normalized', ...
        'Position',[.035 .214 .61 .052],'Visible','off', ...
        'Callback',@quickPaletteChanged);
    ui.editorInputMode='none';
    ui.valueChoiceValues={};
    ui.quickPaletteValues={};
    makeButton(ui.tabEditor,'Apply',[.675 .214 .285 .052],@applyValue);
    ui.openPaletteButton=makeButton(ui.tabEditor,'Choose palette...', ...
        [.035 .145 .46 .052],@pickColor);
    makeButton(ui.tabEditor,'Remove control', ...
        [.515 .145 .445 .052],@removeWorkspace);
    makeButton(ui.tabEditor,'Reset to master', ...
        [.035 .070 .925 .052],@resetProperty);
    makeText(ui.tabEditor, ...
        'Use Properties to add controls. Hide a control without changing its value.', ...
        [.035 .013 .93 .042]);

    makeText(ui.tabProps,'Object:',[.035 .928 .15 .035]);
    ui.group = uicontrol(ui.tabProps,'Style','popupmenu','String', ...
        {'Dataset','Axes','XRuler','YRuler','XLabel','YLabel','Legend','XTicks','YTicks','Panel','Custom'}, ...
        'Units','normalized','Position',[.19 .921 .77 .047], ...
        'Callback',@filterProperties);
    makeText(ui.tabProps,'Search:',[.035 .864 .15 .033]);
    ui.search = uicontrol(ui.tabProps,'Style','edit','String','', ...
        'Units','normalized','Position',[.19 .853 .77 .049], ...
        'HorizontalAlignment','left','Callback',@filterProperties);
    ui.properties = uicontrol(ui.tabProps,'Style','listbox','String',{' '}, ...
        'Min',0,'Max',1,'Units','normalized', ...
        'Position',[.035 .405 .93 .421]);
    makeButton(ui.tabProps,'Add selected to My Editor  >>', ...
        [.035 .337 .93 .055], @addWorkspace);
    makeText(ui.tabProps,'Master-style notes (saved with the master):', ...
        [.035 .286 .93 .038]);
    ui.masterNotes = uicontrol(ui.tabProps,'Style','edit', ...
        'String',S.masterNotes,'Max',20,'Min',0, ...
        'Units','normalized','Position',[.035 .082 .93 .202], ...
        'HorizontalAlignment','left','Callback',@masterNotesEdited, ...
        'TooltipString','Document plot conventions, fonts, intended journal, or palette usage');
    makeText(ui.tabProps, ...
        'Search native properties, then add controls to My Editor.', ...
        [.035 .020 .93 .045]);

    makeText(ui.tabSeries,'Imported datasets:',[.035 .928 .7 .035]);
    ui.seriesList = uicontrol(ui.tabSeries,'Style','listbox', ...
        'String',{'(none)'},'Units','normalized', ...
        'Position',[.035 .42 .93 .50], 'Callback',@seriesChanged);
    makeButton(ui.tabSeries,'Remove selected series', ...
        [.035 .341 .93 .058], @removeSeries);
    makeText(ui.tabSeries,'Series name:',[.035 .28 .28 .038]);
    ui.seriesName = uicontrol(ui.tabSeries,'Style','edit', ...
        'Units','normalized','Position',[.31 .268 .65 .056], ...
        'HorizontalAlignment','left','String','', ...
        'Callback',@renameSeries);
    ui.legendCheck = uicontrol(ui.tabSeries,'Style','checkbox', ...
        'String','Display legend', 'Units','normalized', ...
        'Position',[.035 .188 .60 .055], 'Value',0, ...
        'Callback',@legendChanged);
    makeText(ui.tabSeries,'Select a dataset to edit its marker and line properties.', ...
        [.035 .087 .92 .065]);

    makeText(ui.tabPalette,'CorelDRAW XML palette (RGB/sRGB):', ...
        [.035 .918 .85 .042]);
    makeButton(ui.tabPalette,'Load XML palette...', ...
        [.035 .845 .93 .062], @loadPalette);
    ui.paletteInfo = uicontrol(ui.tabPalette,'Style','text', ...
        'String','Built-in palette','Units','normalized', ...
        'Position',[.035 .783 .93 .045], ...
        'HorizontalAlignment','left');
    % Color preview is a paginated swatch grid, not a monochrome listbox.
    % Color codes are rendered in their actual RGB foreground color.
    ui.palettePrev=makeButton(ui.tabPalette,'< Previous', ...
        [.035 .746 .26 .044],@prevPalettePage);
    ui.palettePageLabel=makeText(ui.tabPalette,'Page 1', ...
        [.365 .749 .30 .035]);
    ui.paletteNext=makeButton(ui.tabPalette,'Next >', ...
        [.704 .746 .26 .044],@nextPalettePage);
    ui.paletteButtons=gobjects(1,12);
    ui.paletteRGB=gobjects(1,12);
    ui.paletteNames=gobjects(1,12);
    for ci=1:12
        rowIndex=floor((ci-1)/3);
        colIndex=mod(ci-1,3);
        xpos=.035+colIndex*.316;
        ypos=.674-rowIndex*.118;
        ui.paletteButtons(ci)=uicontrol(ui.tabPalette,'Style','pushbutton', ...
            'String','', 'Units','normalized', ...
            'Position',[xpos ypos+.037 .058 .052], ...
            'Callback',@swatchClicked);
        ui.paletteRGB(ci)=uicontrol(ui.tabPalette,'Style','text', ...
            'String','', 'Units','normalized', ...
            'Position',[xpos+.063 ypos+.042 .247 .047], ...
            'HorizontalAlignment','left','FontSize',9);
        ui.paletteNames(ci)=uicontrol(ui.tabPalette,'Style','text', ...
            'String','', 'Units','normalized', ...
            'Position',[xpos ypos-.002 .29 .032], ...
            'HorizontalAlignment','left','FontSize',8);
    end
    ui.swatch = uicontrol(ui.tabPalette,'Style','text', ...
        'String','No swatch selected', 'Units','normalized', ...
        'Position',[.035 .218 .93 .046]);
    ui.colorTarget = uicontrol(ui.tabPalette,'Style','popupmenu', ...
        'String',{'Line color','Marker face','Marker edge'}, ...
        'Units','normalized','Position',[.035 .162 .93 .046]);
    makeButton(ui.tabPalette,'Assign to selected series (copy RGB)', ...
        [.035 .088 .93 .056], @assignColor);
    makeText(ui.tabPalette,'Colors are fixed on assignment, never live-linked.', ...
        [.035 .028 .93 .044]);

    makeText(ui.tabPanel,'Panel preset:',[.035 .918 .43 .045]);
    ui.panelPreset = uicontrol(ui.tabPanel,'Style','popupmenu', ...
        'String',{'Normal: 85 mm','Double: 175 mm','Custom'}, ...
        'Units','normalized','Position',[.48 .908 .47 .057], ...
        'Callback',@panelPresetChanged);
    fields = {'widthMM','heightMM','leftMM','rightMM','bottomMM','topMM'};
    captions = {'Plot area width (mm)','Plot area height (mm)', ...
        'Left margin (mm)','Right margin (mm)', ...
        'Bottom margin (mm)','Top margin (mm)'};
    ui.panelEdits = struct();
    for k=1:numel(fields)
        yy=.82-(k-1)*.098;
        makeText(ui.tabPanel,captions{k},[.035 yy .53 .050]);
        ui.panelEdits.(fields{k}) = uicontrol(ui.tabPanel, ...
            'Style','edit','String',num2str(S.panel.(fields{k})), ...
            'Units','normalized','Position',[.60 yy .35 .056], ...
            'Callback',@panelEdited);
    end
    ui.panelPageInfo=makeText(ui.tabPanel,'', [.035 .275 .93 .043]);
    ui.emfCorrection=uicontrol(ui.tabPanel,'Style','checkbox', ...
        'String','CorelDRAW EMF font correction (not PDF)', ...
        'Value',double(S.export.ApplyEMFCorrection), ...
        'Units','normalized','Position',[.035 .222 .93 .045], ...
        'Callback',@exportSettingsEdited);
    makeButton(ui.tabPanel,'Calibrate X/Y EMF font sizes...', ...
        [.035 .151 .93 .055],@openEMFFontCalibration);
    makeText(ui.tabPanel, ...
        'Four independent font corrections are saved in master styles.', ...
        [.035 .106 .93 .039]);


    % All X/Y axes settings, including labels, share a compact editor.
    % Box outline and tick sides are independent. Native MATLAB Box='on'
    % produces opposite-side ticks; a top/right frame-only mode uses two
    % noninteractive vector lines while preserving real lower/left ticks.
    ui.boxCheck=uicontrol(ui.tabTicks,'Style','checkbox', ...
        'String','Box outline','Units','normalized', ...
        'Position',[.045 .025 .29 .049], ...
        'Callback',@boxSettingsEdited);
    makeText(ui.tabTicks,'Box ticks:',[.355 .034 .18 .033]);
    ui.boxTickSides=uicontrol(ui.tabTicks,'Style','popupmenu', ...
        'String',{'Bottom/left only','All four sides'}, ...
        'Units','normalized','Position',[.54 .030 .41 .048], ...
        'Callback',@boxSettingsEdited, ...
        'TooltipString','Draw top/right border without ticks, or use MATLAB native four-sided ticks');
    ui.axesSubtabs=uitabgroup('Parent',ui.tabTicks, ...
        'Units','normalized','Position',[.025 .108 .95 .87]);
    ui.xAxisTab=uitab(ui.axesSubtabs,'Title','X axis');
    ui.yAxisTab=uitab(ui.axesSubtabs,'Title','Y axis');
    ui.limitInputs=struct();ui.tickInputs=struct();
    ui.axisTextInputs=struct();
    createAxisControls('X',ui.xAxisTab);
    createAxisControls('Y',ui.yAxisTab);
    makeText(ui.tabPanel, ...
        'Plot area remains fixed when text changes. Margins determine export canvas.  EMF correction is empirical.', ...
        [.035 .040 .93 .105]);

    ui.status = uicontrol(ui.fig,'Style','text','String','Ready', ...
        'Units','normalized','Position',[.018 .018 .965 .040], ...
        'BackgroundColor',get(ui.fig,'Color'), ...
        'HorizontalAlignment','left','FontSize',9);

    S.catalog = buildCatalog();
    S.workspace = S.workspace(ismember(S.workspace,{S.catalog.key}));
    updateEverything();
    status('Ready. Select an Excel workbook, then paste its =SERIES(...) formula.');


    function createAxisControls(side,tab)
        % Combine typography, limits and major ticks in one tab per axis.
        k=side;t=struct();li=struct();a=struct();
        makeText(tab,[side '-axis label'],[.035 .936 .29 .036]);
        a.label=uicontrol(tab,'Style','edit','String',side, ...
            'Units','normalized','Position',[.33 .927 .635 .049], ...
            'HorizontalAlignment','left');
        makeText(tab,'Label font (pt)',[.035 .867 .31 .037]);
        a.labelFont=uicontrol(tab,'Style','edit','String','10', ...
            'Units','normalized','Position',[.35 .858 .18 .049]);
        makeText(tab,'Tick font (pt)',[.55 .867 .27 .037]);
        a.tickFont=uicontrol(tab,'Style','edit','String','8', ...
            'Units','normalized','Position',[.83 .858 .135 .049]);
        makeText(tab,'Tick-label offset (pt)',[.035 .799 .54 .035]);
        a.tickGap=uicontrol(tab,'Style','edit','String','0', ...
            'Units','normalized','Position',[.67 .788 .295 .048]);
        li.auto=uicontrol(tab,'Style','checkbox','Value',1, ...
            'String','Automatic axis extents','Units','normalized', ...
            'Position',[.035 .731 .80 .048], ...
            'Callback',@limitAutoChanged);
        makeText(tab,'Min',[.035 .671 .10 .033]);
        li.min=uicontrol(tab,'Style','edit','String','', ...
            'Units','normalized','Position',[.135 .662 .32 .049], ...
            'Enable','off');
        makeText(tab,'Max',[.52 .671 .12 .033]);
        li.max=uicontrol(tab,'Style','edit','String','', ...
            'Units','normalized','Position',[.65 .662 .315 .049], ...
            'Enable','off');
        makeText(tab,'Major tick marks',[.035 .604 .43 .036]);
        t.mode=uicontrol(tab,'Style','popupmenu', ...
            'String',{'Auto','Manual'},'Units','normalized', ...
            'Position',[.55 .599 .415 .049]);
        makeText(tab,'Positions (e.g. -2:0.5:2)',[.035 .546 .93 .035]);
        t.positions=uicontrol(tab,'Style','edit','String','', ...
            'Units','normalized','Position',[.035 .496 .93 .048], ...
            'HorizontalAlignment','left');
        makeText(tab,'Tick number format',[.035 .445 .43 .034]);
        t.format=uicontrol(tab,'Style','popupmenu', ...
            'String',{'Auto','Number (fixed)','General','Scientific (e)', ...
                'Engineering (e3)','SI prefixes','Custom sprintf'}, ...
            'Units','normalized','Position',[.48 .438 .485 .048]);
        makeText(tab,'Decimals',[.035 .386 .20 .034]);
        t.decimals=uicontrol(tab,'Style','edit','String','2', ...
            'Units','normalized','Position',[.235 .379 .18 .046]);
        makeText(tab,'Prefix',[.45 .386 .13 .034]);
        t.prefix=uicontrol(tab,'Style','edit','String','', ...
            'Units','normalized','Position',[.58 .379 .135 .046]);
        makeText(tab,'Suffix',[.73 .386 .14 .034]);
        t.suffix=uicontrol(tab,'Style','edit','String','', ...
            'Units','normalized','Position',[.86 .379 .105 .046]);
        makeText(tab,'Custom sprintf pattern',[.035 .325 .49 .034]);
        t.customFormat=uicontrol(tab,'Style','edit','String','%.2f', ...
            'Units','normalized','Position',[.53 .319 .435 .047]);
        makeText(tab,'Explicit labels (use | as separator; optional)', ...
            [.035 .264 .93 .034]);
        t.labels=uicontrol(tab,'Style','edit','String','', ...
            'Units','normalized','Position',[.035 .207 .93 .051], ...
            'HorizontalAlignment','left');
        makeButton(tab,['Apply ' side ' axis settings'], ...
            [.035 .115 .93 .068],@(~,~)applyAxisUI(k));
        makeText(tab,'Axis labels, ticks and limits saved per figure.', ...
            [.035 .038 .93 .05]);
        ui.limitInputs.(k)=li;ui.tickInputs.(k)=t;
        ui.axisTextInputs.(k)=a;
    end

    function syncAxisTypographyControls()
        for side={'X','Y'}
            k=side{1}; c=ui.axisTextInputs.(k);
            labelGroup=[k 'Label'];rulerGroup=[k 'Ruler'];
            txt=S.overrides.(labelGroup);
            if isfield(txt,'String'),v=txt.String;else,v=k;end
            set(c.label,'String',v);
            if isfield(txt,'FontSize'),v=txt.FontSize;
            else,v=get(ui.ax.([k 'Label']),'FontSize');end
            set(c.labelFont,'String',num2str(v));
            r=S.overrides.(rulerGroup);
            if isfield(r,'FontSize'),v=r.FontSize;
            else,v=get(ui.ax.([k 'Axis']),'FontSize');end
            set(c.tickFont,'String',num2str(v));
            set(c.tickGap,'String',num2str(S.custom.([k 'TickGapPt'])));
        end
    end

    function applyAxisUI(k)
        % Validate all inputs before committing any of them.
        c=ui.axisTextInputs.(k);li=ui.limitInputs.(k);
        tc=ui.tickInputs.(k);
        try
            labelText=get(c.label,'String');
            labelFont=str2double(get(c.labelFont,'String'));
            tickFont=str2double(get(c.tickFont,'String'));
            gap=str2double(get(c.tickGap,'String'));
            if any(~isfinite([labelFont tickFont gap])) || ...
                    labelFont<=0 || tickFont<=0
                error('Font sizes must be positive and gap must be finite.');
            end
            tick=S.ticks.(k);
            modes={'auto','manual'};tick.Mode=modes{get(tc.mode,'Value')};
            tick.Positions=parseTickSequence(get(tc.positions,'String'));
            if strcmp(tick.Mode,'manual') && isempty(tick.Positions)
                error('Enter major tick positions for Manual mode.');
            end
            fmts={'auto','number','general','scientific','engineering','si','custom'};
            tick.Format=fmts{get(tc.format,'Value')};
            tick.Decimals=str2double(get(tc.decimals,'String'));
            if ~isfinite(tick.Decimals)||tick.Decimals<0|| ...
                    tick.Decimals>12||fix(tick.Decimals)~=tick.Decimals
                error('Decimals must be an integer from 0 to 12.');
            end
            tick.Prefix=get(tc.prefix,'String');
            tick.Suffix=get(tc.suffix,'String');
            tick.CustomFormat=get(tc.customFormat,'String');
            if strcmp(tick.Format,'custom')
                validateFormat(tick.CustomFormat);
            end
            tick.Labels=get(tc.labels,'String');
            if ~isempty(tick.Labels)
                if ~strcmp(tick.Mode,'manual')
                    error('Explicit labels require Manual major ticks.');
                end
                labels=strsplit(tick.Labels,'|');
                if numel(labels)~=numel(tick.Positions)
                    error('Provide one explicit label per major tick.');
                end
            end
            auto=get(li.auto,'Value')~=0;
            if ~auto
                low=str2double(get(li.min,'String'));
                high=str2double(get(li.max,'String'));
                if ~isfinite(low)||~isfinite(high)||low>=high
                    error('%s-axis minimum must be smaller than maximum.',k);
                end
                if strcmpi(get(ui.ax,[k 'Scale']),'log') && low<=0
                    error('A logarithmic axis needs positive limits.');
                end
            end
            S.overrides.([k 'Label']).String=labelText;
            S.overrides.([k 'Label']).FontSize=labelFont;
            S.overrides.([k 'Ruler']).FontSize=tickFont;
            S.custom.([k 'TickGapPt'])=gap;
            if auto
                S.overrides.Axes.([k 'LimMode'])='auto';
                if isfield(S.overrides.Axes,[k 'Lim'])
                    S.overrides.Axes=rmfield(S.overrides.Axes,[k 'Lim']);
                end
            else
                S.overrides.Axes.([k 'Lim'])=[low high];
                S.overrides.Axes.([k 'LimMode'])='manual';
            end
            S.ticks.(k)=tick;
            redraw();syncLimitControls();syncTickControls();
            status([k '-axis settings applied.']);
        catch err
            errordlg(err.message,'Invalid axis settings');
        end
    end

    function syncLimitControls()
        for side={'X','Y'}
            k=side{1};lo=S.overrides.Axes;
            limProp=[k 'Lim'];modeProp=[k 'LimMode'];
            manual=isfield(lo,limProp) && ...
                (~isfield(lo,modeProp) || strcmpi(lo.(modeProp),'manual'));
            set(ui.limitInputs.(k).auto,'Value',double(~manual));
            if manual
                limits=lo.(limProp);
                set(ui.limitInputs.(k).min,'String',num2str(limits(1),12));
                set(ui.limitInputs.(k).max,'String',num2str(limits(2),12));
            else
                set(ui.limitInputs.(k).min,'String','');
                set(ui.limitInputs.(k).max,'String','');
            end
            set(ui.limitInputs.(k).min,'Enable',onOff(manual));
            set(ui.limitInputs.(k).max,'Enable',onOff(manual));
        end
    end

    function limitAutoChanged(~,~)
        for side={'X','Y'}
            k=side{1};manual=~logical(get(ui.limitInputs.(k).auto,'Value'));
            set(ui.limitInputs.(k).min,'Enable',onOff(manual));
            set(ui.limitInputs.(k).max,'Enable',onOff(manual));
            if manual && isempty(get(ui.limitInputs.(k).min,'String'))
                lim=get(ui.ax,[k 'Lim']);
                set(ui.limitInputs.(k).min,'String',num2str(lim(1),12));
                set(ui.limitInputs.(k).max,'String',num2str(lim(2),12));
            end
        end
    end

    function syncTickControls()
        styles={'auto','number','general','scientific','engineering','si','custom'};
        for side={'X','Y'}
            k=side{1};t=S.ticks.(k);c=ui.tickInputs.(k);
            set(c.mode,'Value',1+strcmpi(t.Mode,'manual'));
            set(c.positions,'String',tickPositionString(t.Positions));
            fmt=find(strcmp(styles,lower(t.Format)),1);
            if isempty(fmt),fmt=1;end
            set(c.format,'Value',fmt);
            set(c.decimals,'String',num2str(t.Decimals));
            set(c.prefix,'String',t.Prefix);
            set(c.suffix,'String',t.Suffix);
            set(c.customFormat,'String',t.CustomFormat);
            set(c.labels,'String',t.Labels);
        end
    end

    function applyTickSettings(ax,side)
        t=S.ticks.(side);
        tickProp=[side 'Tick'];labelProp=[side 'TickLabel'];
        modeProp=[side 'TickMode'];labelModeProp=[side 'TickLabelMode'];
        try
            if strcmpi(t.Mode,'manual') && ~isempty(t.Positions)
                set(ax,tickProp,t.Positions);
            else
                set(ax,modeProp,'auto');
            end
            if strcmpi(t.Format,'auto') && isempty(t.Prefix) && ...
                    isempty(t.Suffix) && isempty(t.Labels)
                set(ax,labelModeProp,'auto');
                return;
            end
            vals=get(ax,tickProp);
            if isempty(vals),return;end
            if ~isempty(t.Labels)
                labels=strsplit(t.Labels,'|');
                if numel(labels)~=numel(vals)
                    error('%s tick labels: %d labels for %d ticks.', ...
                        side,numel(labels),numel(vals));
                end
                labels=cellfun(@strtrim,labels,'UniformOutput',false);
            else
                labels=cell(1,numel(vals));
                for q=1:numel(vals)
                    labels{q}=[t.Prefix formatTickNumber(vals(q),t,vals) t.Suffix];
                end
            end
            set(ax,labelProp,labels);
        catch err
            status(['Tick formatting: ' err.message]);
        end
    end

    function vals=parseTickSequence(raw)
        if isnumeric(raw),vals=double(raw(:)');else
            if isstring(raw),raw=char(raw);end
            raw=strtrim(raw);
            if isempty(raw)||strcmp(raw,'[]'),vals=[];return;end
            number='[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?';
            pattern=['^\s*(' number ')\s*:\s*(' number ')\s*:\s*(' number ')\s*$'];
            tok=regexp(raw,pattern,'tokens','once');
            if ~isempty(tok)
                abc=cellfun(@str2double,tok);
                if ~all(isfinite(abc)) || abc(2)<=0
                    error('Tick step must be a positive finite number.');
                end
                vals=abc(1):abc(2):abc(3);
            else
                raw=regexprep(raw,'[\[\],;]',' ');
                pieces=regexp(strtrim(raw),'\s+','split');
                vals=cellfun(@str2double,pieces);
            end
        end
        if numel(vals)>500 || any(~isfinite(vals)) || any(diff(vals)<=0)
            error('Tick positions must be finite, strictly increasing, and at most 500 values.');
        end
    end

    function str=tickPositionString(vals)
        if isempty(vals),str='';else,str=mat2str(vals,8);end
    end

    function validateFormat(f)
        if isstring(f),f=char(f);end
        % Exactly one safe numeric sprintf conversion; no arbitrary %n/%s.
        if isempty(regexp(f,'^([^%]|%%)*%[-+ 0#]*[0-9]{0,2}(?:\.[0-9]{1,2})?[eEfFgG]([^%]|%%)*$', 'once'))
            error('Custom format needs one numeric conversion, e.g. %.2f, %+.3e or %.4g.');
        end
    end

    function str=formatTickNumber(x,t,allValues)
        d=t.Decimals;
        switch lower(t.Format)
            case 'number',str=sprintf(['%.' num2str(d) 'f'],x);
            case 'scientific',str=sprintf(['%.' num2str(d) 'e'],x);
            case 'general',str=sprintf(['%.' num2str(max(1,d)) 'g'],x);
            case 'custom',str=sprintf(t.CustomFormat,x);
            case {'engineering','si'}
                scale=max(abs(allValues));
                if scale==0,e=0;else,e=3*floor(log10(scale)/3);end
                e=max(-24,min(24,e));
                scaled=x/10^e;
                value=sprintf(['%.' num2str(d) 'f'],scaled);
                if strcmpi(t.Format,'si')
                    exponents=-24:3:24;
                    prefixes={'y','z','a','f','p','n','u','m','','k','M','G','T','P','E','Z','Y'};
                    ind=find(exponents==e,1);
                    str=[value ' ' prefixes{ind}];
                elseif e==0
                    str=value;
                else
                    str=sprintf('%se%+03d',value,e);
                end
            otherwise,str=sprintf('%.6g',x);
        end
    end

    function bg=contrastBG(rgb)
        % Keep color-accurate foreground labels legible on both light and
        % dark colors without changing the selected RGB value.
        rgb=double(rgb(:)');
        linear=zeros(1,3);
        for c=1:3
            if rgb(c)<=.04045,linear(c)=rgb(c)/12.92;
            else,linear(c)=((rgb(c)+.055)/1.055)^2.4;end
        end
        lum=sum(linear.*[.2126 .7152 .0722]);
        if lum>.23,bg=[.1 .1 .1];else,bg=[.97 .97 .97];end
    end

    function state=onOff(b)
        if b,state='on';else,state='off';end
    end
    function label=onSelection(b)
        if b,label='*';else,label=' ';end
    end

    % -------------------- UI / state functions --------------------
    function status(message)
        S.lastStatus = char(message);
        if isgraphics(ui.status), set(ui.status,'String',S.lastStatus); end
        drawnow limitrate;
    end

    function updateEverything()
        set(ui.workbook,'String',S.workbook);
        syncSeriesControls();
        syncPaletteControls();
        syncPanelControls();
        syncExportControls();
        syncMasterControls();
        syncLimitControls();
        syncTickControls();
        syncBoxControls();
        redraw();
        syncAxisTypographyControls();
        filterProperties();
    end

    function syncSeriesControls()
        names = {S.datasets.name};
        if isempty(names), names = {'(No series)'}; end
        S.activeSeries = min(max(S.activeSeries,1),max(1,numel(S.datasets)));
        set(ui.seriesList,'String',names,'Value',S.activeSeries);
        set(ui.editorSeries,'String',names,'Value',S.activeSeries);
        if isempty(S.datasets)
            set(ui.seriesName,'String','');
        else
            set(ui.seriesName,'String',S.datasets(S.activeSeries).name);
        end
        set(ui.legendCheck,'Value',double(S.legendOn));
    end

    function masterNameEdited(~,~)
        name=strtrim(get(ui.masterName,'String'));
        if isempty(name)
            set(ui.masterName,'String',S.masterName);
            return;
        end
        S.masterName=name;
    end

    function masterNotesEdited(~,~)
        S.masterNotes=notesAsText(get(ui.masterNotes,'String'));
    end

    function syncMasterControls()
        set(ui.masterName,'String',S.masterName);
        set(ui.masterNotes,'String',S.masterNotes);
    end

    function syncPanelControls()
        for j=1:numel(fields)
            set(ui.panelEdits.(fields{j}),'String',sprintf('%.5g',S.panel.(fields{j})));
        end
        switch lower(S.panel.preset)
            case 'normal', index=1;
            case 'double', index=2;
            otherwise, index=3;
        end
        set(ui.panelPreset,'Value',index);
        dims=panelCanvasMM();
        set(ui.panelPageInfo,'String',sprintf( ...
            'Fixed plot box: %.3g x %.3g mm   |   Export canvas: %.3g x %.3g mm', ...
            S.panel.widthMM,S.panel.heightMM,dims(1),dims(2)));
    end

    function syncExportControls()
        set(ui.emfCorrection,'Value',double(S.export.ApplyEMFCorrection));
    end

    function syncBoxControls()
        % Read the requested setting, not the preview's emulated Box=off.
        showBox=false;
        if isfield(S.overrides.Axes,'Box')
            showBox=isOnValue(S.overrides.Axes.Box);
        end
        set(ui.boxCheck,'Value',double(showBox));
        if strcmp(S.custom.BoxTickSides,'all')
            set(ui.boxTickSides,'Value',2);
        else
            set(ui.boxTickSides,'Value',1);
        end
    end

    function boxSettingsEdited(~,~)
        S.overrides.Axes.Box=onOff(get(ui.boxCheck,'Value')~=0);
        modes={'primary','all'};
        S.custom.BoxTickSides=modes{get(ui.boxTickSides,'Value')};
        redraw();
    end

    function bool=isOnValue(val)
        % Accept char/string/on-off enumeration/logical saved by all versions.
        if islogical(val) || isnumeric(val)
            bool=logical(val(1));
        else
            bool=strcmpi(char(string(val)),'on');
        end
    end

    function exportSettingsEdited(~,~)
        S.export.ApplyEMFCorrection=logical(get(ui.emfCorrection,'Value'));
    end

    function captureVisibleFontSizes()
        % Text entered in either X/Y axis tab must be saved even if the
        % user has not pressed its separate "Apply axis settings" button.
        for side={'X','Y'}
            k=side{1};ctl=ui.axisTextInputs.(k);
            labelSize=str2double(get(ctl.labelFont,'String'));
            tickSize=str2double(get(ctl.tickFont,'String'));
            if ~isfinite(labelSize) || labelSize<=0 || ...
                    ~isfinite(tickSize) || tickSize<=0
                error('%s-axis font sizes must be positive numbers.',k);
            end
            S.overrides.([k 'Label']).FontSize=labelSize;
            S.overrides.([k 'Ruler']).FontSize=tickSize;
        end
    end

    function openEMFFontCalibration(~,~)
        % The user enters actual point sizes observed after importing the
        % most recently exported EMF into CorelDRAW. Each of four scaling
        % factors is updated independently, then retained in the master.
        % A calibration can also be edited directly, without a test import.
        labels={'X tick labels','Y tick labels','X axis label','Y axis label'};
        keys={'XTick','YTick','XLabel','YLabel'};
        targets=zeros(1,4);
        for n=1:2
            k='X';if n==2,k='Y';end
            ctl=ui.axisTextInputs.(k);
            targets(n)=str2double(get(ctl.tickFont,'String'));
            targets(n+2)=str2double(get(ctl.labelFont,'String'));
        end
        if any(~isfinite(targets)) || any(targets<=0)
            errordlg('Enter valid target font sizes in Axes & ticks first.', ...
                'Font calibration');return;
        end
        d=dialog('Name','CorelDRAW EMF font-size calibration', ...
            'Position',[160 160 770 405],'WindowStyle','normal','Resize','off');
        uicontrol('Parent',d,'Style','text','Units','pixels', ...
            'Position',[20 359 735 30],'HorizontalAlignment','left', ...
            'String',['Enter sizes measured in CorelDRAW for the current ' ...
                'EMF, or edit the correction factor directly.']);
        headers={'Text object','Target in MATLAB (pt)','Observed in CorelDRAW (pt)','EMF multiplier'};
        x=[22 205 409 612];w=[175 194 199 134];
        for q=1:4
            uicontrol('Parent',d,'Style','text','String',headers{q}, ...
                'FontWeight','bold','Units','pixels', ...
                'Position',[x(q) 317 w(q) 28], ...
                'HorizontalAlignment','left');
        end
        factorFields=gobjects(1,4);observedFields=gobjects(1,4);
        for j=1:4
            y=311-j*48;
            uicontrol('Parent',d,'Style','text', ...
                'String',labels{j},'Units','pixels', ...
                'Position',[x(1) y w(1) 25], ...
                'HorizontalAlignment','left');
            uicontrol('Parent',d,'Style','text', ...
                'String',sprintf('%.5g',targets(j)), ...
                'Units','pixels','Position',[x(2) y w(2) 25], ...
                'HorizontalAlignment','left');
            observedFields(j)=uicontrol('Parent',d,'Style','edit', ...
                'String','','Units','pixels', ...
                'Position',[x(3) y 160 28], ...
                'TooltipString','CorelDRAW reported point size of the last exported EMF');
            factorFields(j)=uicontrol('Parent',d,'Style','edit', ...
                'String',sprintf('%.9g',S.export.FontFactors.(keys{j})), ...
                'Units','pixels','Position',[x(4) y 118 28]);
        end
        uicontrol('Parent',d,'Style','text','Units','pixels', ...
            'Position',[22 77 720 38],'HorizontalAlignment','left', ...
            'String',['Recalibrate: new multiplier = previous multiplier ' ...
                'x target size / measured size. Blank measurements are ignored.']);
        uicontrol('Parent',d,'Style','pushbutton', ...
            'String','Calculate from measurements', ...
            'Units','pixels','Position',[25 22 256 40], ...
            'Callback',@calculateFromMeasurements);
        uicontrol('Parent',d,'Style','pushbutton', ...
            'String','Apply corrections','Units','pixels', ...
            'Position',[448 22 155 40],'Callback',@applyFactors);
        uicontrol('Parent',d,'Style','pushbutton', ...
            'String','Cancel','Units','pixels', ...
            'Position',[616 22 125 40],'Callback',@(~,~)delete(d));

        function calculateFromMeasurements(~,~)
            try
                for m=1:4
                    measured=strtrim(get(observedFields(m),'String'));
                    if isempty(measured),continue;end
                    value=str2double(measured);
                    factor=str2double(get(factorFields(m),'String'));
                    if ~isfinite(value)||value<=0|| ...
                            ~isfinite(factor)||factor<=0
                        error('%s: measured size and factor must be positive.',labels{m});
                    end
                    set(factorFields(m),'String',sprintf('%.9g', ...
                        factor*targets(m)/value));
                end
            catch err
                errordlg(err.message,'Invalid calibration');
            end
        end

        function applyFactors(~,~)
            try
                next=S.export.FontFactors;
                for m=1:4
                    value=str2double(get(factorFields(m),'String'));
                    if ~isfinite(value)||value<=0||value>10
                        error('%s correction must be a number between 0 and 10.',labels{m});
                    end
                    next.(keys{m})=value;
                end
                captureVisibleFontSizes();
                S.export.FontFactors=next;
                delete(d);
                status('Four independent EMF font corrections updated. Save master to reuse.');
            catch err
                errordlg(err.message,'Invalid calibration');
            end
        end
    end

    function syncPaletteControls()
        entries=S.palette.colors;
        if isempty(entries)
            set(ui.paletteInfo,'String','No RGB/sRGB swatches available');
            if strcmp(ui.editorInputMode,'palette')
                updateEditor();
            end
            for j=1:12
                set([ui.paletteButtons(j),ui.paletteRGB(j),ui.paletteNames(j)], ...
                    'Visible','off');
            end
            return;
        end
        S.palettePage=min(max(1,S.palettePage),ceil(numel(entries)/12));
        S.paletteSelected=min(max(1,S.paletteSelected),numel(entries));
        set(ui.paletteInfo,'String',sprintf('%s | %d fixed swatches', ...
            S.palette.name,numel(entries)));
        pages=ceil(numel(entries)/12);
        set(ui.palettePageLabel,'String',sprintf('Page %d / %d',S.palettePage,pages));
        set(ui.palettePrev,'Enable',onOff(S.palettePage>1));
        set(ui.paletteNext,'Enable',onOff(S.palettePage<pages));
        for j=1:12
            ix=(S.palettePage-1)*12+j;
            if ix<=numel(entries)
                c=entries(ix);
                set(ui.paletteButtons(j),'Visible','on', ...
                    'BackgroundColor',c.rgb,'UserData',ix, ...
                    'String',onSelection(ix==S.paletteSelected));
                set(ui.paletteRGB(j),'Visible','on', ...
                    'String',sprintf('%d,%d,%d',round(c.rgb*255)), ...
                    'ForegroundColor',c.rgb, ...
                    'BackgroundColor',contrastBG(c.rgb), ...
                    'TooltipString',rgbHex(c.rgb));
                set(ui.paletteNames(j),'Visible','on','String',c.name, ...
                    'TooltipString',c.name);
            else
                set([ui.paletteButtons(j),ui.paletteRGB(j),ui.paletteNames(j)], ...
                    'Visible','off');
            end
        end
        c=entries(S.paletteSelected);
        set(ui.swatch,'BackgroundColor',contrastBG(c.rgb), ...
            'ForegroundColor',c.rgb, ...
            'String',sprintf('Selected: %s  %s',c.name,rgbHex(c.rgb)));
    end

    function redraw()
        try
            S.plotHandles = drawPlot(ui.ax,false);
        catch err
            status(['Plot error: ' err.message]);
        end
        updateEditor();
        syncBoxControls();
    end

    function hs = drawPlot(ax,isExport)
        cla(ax,'reset');
        hold(ax,'on');
        set(ax,'Color',[1 1 1],'Box','off');
        setInnerPositionConstraint(ax);
        if isExport
            set(ax,'Units','centimeters', 'Position',axesCM());
        else
            set(ax,'Units','normalized', 'Position',axesNormalized());
        end
        hs=cell(1,numel(S.datasets));
        for j=1:numel(S.datasets)
            d=S.datasets(j);
            if isempty(d.x) || isempty(d.y), continue; end
            opts={'DisplayName',d.name};
            h=plot(ax,d.x(:),d.y(:),opts{:});
            applyStyle(h,d.style);
            hs{j}=h;
        end
        hold(ax,'off');
        if isempty(S.datasets)
            text(ax,.5,.5,'Paste an Excel =SERIES(...) formula to begin', ...
                'Units','normalized','HorizontalAlignment','center', ...
                'Color',[.55 .55 .55],'FontSize',11);
        end
        % Establish default label strings, then native per-figure overrides.
        xlabel(ax,'X'); ylabel(ax,'Y');
        applyStored(ax,'Axes',ax);
        applyStored(ax,'XRuler',ax.XAxis);
        applyStored(ax,'YRuler',ax.YAxis);
        applyStored(ax,'XLabel',ax.XLabel);
        applyStored(ax,'YLabel',ax.YLabel);
        % Dedicated axis settings are authoritative when applying a plot.
        % They can be edited directly or through the Properties library.
        applyTickSettings(ax,'X');
        applyTickSettings(ax,'Y');
        % Experimental tick-gap setting (undocumented on some releases).
        for side={'X','Y'}
            key=[side{1} 'TickGapPt'];
            ruler=ax.([side{1} 'Axis']);
            if isprop(ruler,'TickLabelGapOffset')
                try set(ruler,'TickLabelGapOffset',S.custom.(key)); catch, end
            end
        end
        if S.legendOn && ~isempty(S.datasets)
            hsValid=[hs{~cellfun(@isempty,hs)}];
            if ~isempty(hsValid)
                lg=legend(ax,hsValid,'Location','best');
                applyStored(ax,'Legend',lg);
            end
        else
            legend(ax,'off');
        end
        % Decorations must not shrink the plot box, even during editing.
        setInnerPositionConstraint(ax);
        if isExport
            set(ax,'Units','centimeters','Position',axesCM());
        else
            set(ax,'Units','normalized','Position',axesNormalized());
        end
        if strcmp(S.custom.BoxTickSides,'primary') && isOnValue(ax.Box)
            addFrameWithoutOppositeTicks(ax);
        end
    end

    function addFrameWithoutOppositeTicks(ax)
        % Box='on' in MATLAB shows a second set of ticks at the top/right.
        % Draw ONLY the missing two borders as vector line objects instead.
        % Manual limits prevent these decorative lines from padding autoscaled
        % plot bounds; redraw recalculates auto limits from data each time.
        xl=ax.XLim; yl=ax.YLim;
        topY=yl(2);rightX=xl(2);
        if strcmpi(char(string(ax.YDir)),'reverse'),topY=yl(1);end
        if strcmpi(char(string(ax.XDir)),'reverse'),rightX=xl(1);end
        xcolor=ax.XColor; ycolor=ax.YColor; lineWidth=ax.LineWidth;
        ax.Box='off';
        ax.XLim=xl;ax.YLim=yl;
        line(ax,xl,[topY topY], ...
            'Color',xcolor,'LineWidth',lineWidth,'LineStyle','-', ...
            'HandleVisibility','off','HitTest','off','Tag','PPEFrameTop');
        line(ax,[rightX rightX],yl, ...
            'Color',ycolor,'LineWidth',lineWidth,'LineStyle','-', ...
            'HandleVisibility','off','HitTest','off','Tag','PPEFrameRight');
    end

    function dims=panelCanvasMM()
        p=S.panel;
        dims=[p.widthMM+p.leftMM+p.rightMM, ...
              p.heightMM+p.bottomMM+p.topMM];
    end

    function pos=axesNormalized()
        p=S.panel;dims=panelCanvasMM();
        pos=[p.leftMM/dims(1),p.bottomMM/dims(2), ...
            p.widthMM/dims(1),p.heightMM/dims(2)];
    end

    function pos=axesCM()
        p=S.panel;
        pos=[p.leftMM,p.bottomMM,p.widthMM,p.heightMM]/10;
    end

    function setInnerPositionConstraint(ax)
        % Preserve the plot box (as distinct from labels/TightInset).
        % R2020a+ uses PositionConstraint; older axes had this legacy name.
        if isprop(ax,'PositionConstraint')
            set(ax,'PositionConstraint','innerposition');
        elseif isprop(ax,'ActivePositionProperty')
            set(ax,'ActivePositionProperty','position');
        end
    end

    function applyStored(~,group,target)
        fieldsNative=fieldnames(S.overrides.(group));
        for kk=1:numel(fieldsNative)
            try
                set(target,fieldsNative{kk},S.overrides.(group).(fieldsNative{kk}));
            catch
                % Properties unavailable in another MATLAB release remain
                % saved in the recipe. They do not prevent displaying data.
            end
        end
    end

    function applyStyle(h,sty)
        % MarkerEdgeColor='none' is accepted by some older MATLAB versions,
        % but not by every newer Line object. Preserve the semantic setting
        % in the recipe and resolve it for each newly created graphics object.
        fns=fieldnames(sty);
        for kk=1:numel(fns)
            prop=fns{kk};
            if strcmp(prop,'MarkerEdgeColor'),continue;end
            try set(h,prop,sty.(prop)); catch, end
        end
        if isfield(sty,'MarkerEdgeColor')
            try applyMarkerEdgeColor(h,sty.MarkerEdgeColor,sty); catch, end
        end
    end

    function applyMarkerEdgeColor(h,value,sty)
        % Prefer genuine 'none' when supported. Otherwise draw the marker
        % edge in the marker face color, or the axes background when hollow.
        % This hides the edge visually while retaining 'none' in the recipe.
        if (ischar(value)||isstring(value)) && ...
                strcmpi(strtrim(char(value)),'none')
            try
                set(h,'MarkerEdgeColor','none');
            catch
                set(h,'MarkerEdgeColor',invisibleEdgeFallback(h,sty));
            end
        else
            set(h,'MarkerEdgeColor',value);
        end
    end

    function c=invisibleEdgeFallback(h,sty)
        face='none';
        if isfield(sty,'MarkerFaceColor')
            face=sty.MarkerFaceColor;
        end
        if isnumeric(face) && numel(face)==3
            c=double(face(:)');
            return;
        end
        if isstring(face),face=char(face);end
        if ischar(face) && ~any(strcmpi(face,{'none','auto'}))
            c=face; % MATLAB named color or #RRGGBB
            return;
        end
        ax=ancestor(h,'axes');
        if ~isempty(ax)
            c=ax.Color;
            if isnumeric(c)&&numel(c)==3,return;end
            fig=ancestor(ax,'figure');
            if ~isempty(fig) && isnumeric(fig.Color) && numel(fig.Color)==3
                c=fig.Color;
                return;
            end
        end
        c=[1 1 1];
    end

    function workbookEdited(~,~)
        S.workbook=strtrim(get(ui.workbook,'String'));
    end

    function browseWorkbook(~,~)
        [f,p]=uigetfile({'*.xlsx;*.xls;*.xlsm','Excel workbooks'}, ...
            'Choose workbook');
        if isequal(f,0),return;end
        S.workbook=fullfile(p,f);
        set(ui.workbook,'String',S.workbook);
        status(['Workbook selected: ' f]);
    end

    function addSeries(~,~)
        workbookEdited();
        if isempty(S.workbook) || ~exist(S.workbook,'file')
            errordlg('Select an existing Excel workbook first.','No workbook');return;
        end
        raw=get(ui.formula,'String');
        if ischar(raw), lines=regexp(raw,'\r?\n','split');
        elseif iscell(raw),lines=raw;else,lines=cellstr(raw);end
        count=0;
        for ii=1:numel(lines)
            formula=strtrim(lines{ii});
            if isempty(formula),continue;end
            try
                d=seriesFromFormula(formula,S.workbook);
                d.style=styleForSeries(numel(S.datasets)+1);
                d.sourceColorName='';
                S.datasets(end+1)=d;
                count=count+1;
            catch err
                errordlg(sprintf('Could not import formula %d:\n%s\n\n%s', ...
                    ii,formula,err.message),'Excel import error');
            end
        end
        if count>0
            S.activeSeries=numel(S.datasets);
            set(ui.formula,'String','');
            syncSeriesControls();redraw();
            status(sprintf('Imported %d Excel series.',count));
        end
    end

    function d=seriesFromFormula(formula,book)
        parts=parseSeries(formula);
        [xs,xr] = splitReference(parts{2});
        [ys,yr] = splitReference(parts{3});
        if isempty(xs) && isempty(ys)
            xs='Sheet1';ys='Sheet1';
        elseif isempty(xs),xs=ys;
        elseif isempty(ys),ys=xs;
        end
        x=readmatrix(book,'Sheet',xs,'Range',xr);
        y=readmatrix(book,'Sheet',ys,'Range',yr);
        if ~isnumeric(x) || ~isnumeric(y) || isempty(x) || isempty(y)
            error('The selected ranges contain no numeric data.');
        end
        if ~isvector(x) || ~isvector(y)
            error('Use one-dimensional Excel ranges for X and Y.');
        end
        x=x(:); y=y(:);
        if numel(x)~=numel(y)
            error('The X range contains %d values and Y contains %d.',numel(x),numel(y));
        end
        name=strtrim(parts{1});
        if numel(name)>=2 && name(1)=='"' && name(end)=='"'
            name=strrep(name(2:end-1),'""','"');
        end
        if isempty(name),name=sprintf('Series %d',numel(S.datasets)+1);end
        order=str2double(parts{4});
        if isnan(order),order=numel(S.datasets)+1;end
        d=struct('name',name,'formula',formula,'workbook',book, ...
            'xSheet',xs,'xRange',xr,'ySheet',ys,'yRange',yr, ...
            'order',order,'x',x,'y',y,'style',struct(), ...
            'sourceColorName','');
    end

    function parts=parseSeries(formula)
        formula=strtrim(formula);
        mt=regexp(formula,'^\s*=?\s*SERIES\s*\((.*)\)\s*$','tokens','once','ignorecase');
        if isempty(mt)
            error('Expected =SERIES(name, x_range, y_range, order).');
        end
        input=mt{1};parts={};start=1;
        inSingle=false;inDouble=false;paren=0;
        k=1;
        while k<=numel(input)
            ch=input(k);
            if ch=='''' && ~inDouble
                if inSingle && k<numel(input) && input(k+1)==''''
                    k=k+1;
                else
                    inSingle=~inSingle;
                end
            elseif ch=='"' && ~inSingle
                if inDouble && k<numel(input) && input(k+1)=='"'
                    k=k+1;
                else
                    inDouble=~inDouble;
                end
            elseif ~inSingle && ~inDouble
                if ch=='(',paren=paren+1;
                elseif ch==')',paren=paren-1;
                elseif ch==',' && paren==0
                    parts{end+1}=strtrim(input(start:k-1)); %#ok<AGROW>
                    start=k+1;
                end
            end
            k=k+1;
        end
        parts{end+1}=strtrim(input(start:end));
        if numel(parts)~=4
            error('Expected four SERIES arguments; found %d.',numel(parts));
        end
    end

    function [sheet,range]=splitReference(ref)
        ref=strtrim(ref);
        exclam=find(ref=='!',1,'last');
        if isempty(exclam)
            sheet='';range=ref;
        else
            sheet=strtrim(ref(1:exclam-1));
            range=strtrim(ref(exclam+1:end));
        end
        if numel(sheet)>=2 && sheet(1)=='''' && sheet(end)==''''
            sheet=sheet(2:end-1);
            sheet=strrep(sheet,'''''','''');
        end
        if contains(sheet,'[') && contains(sheet,']')
            % Excel can qualify references with a workbook. The app uses
            % the selected workbook and extracts only the sheet name.
            sheet=regexprep(sheet,'^.*\]','');
        end
        range=strrep(range,'$','');
        if isempty(regexp(range,'^[A-Za-z]+[0-9]+(:[A-Za-z]+[0-9]+)?$', 'once'))
            error('Unsupported Excel cell range: %s',ref);
        end
    end

    function refreshExcel(~,~)
        count=0;failures={};
        for ii=1:numel(S.datasets)
            d=S.datasets(ii);
            if isempty(d.formula),continue;end
            try
                fresh=seriesFromFormula(d.formula,d.workbook);
                S.datasets(ii).x=fresh.x;
                S.datasets(ii).y=fresh.y;
                count=count+1;
            catch err
                failures{end+1}=sprintf('%s: %s',d.name,err.message); %#ok<AGROW>
            end
        end
        redraw();
        if isempty(failures)
            status(sprintf('Refreshed %d data series from Excel.',count));
        else
            warndlg(strjoin(failures,newline),'Some sources unavailable');
            status(sprintf('Refreshed %d series; %d unavailable.',count,numel(failures)));
        end
    end

    function sty=styleForSeries(n)
        sty=S.styleDefaults;
        colors=S.palette.colors;
        if ~isempty(colors)
            sty.Color=colors(mod(n-1,numel(colors))+1).rgb;
        end
    end

    function editorSeriesChanged(~,~)
        S.activeSeries=get(ui.editorSeries,'Value');
        syncSeriesControls();updateEditor();
    end
    function seriesChanged(~,~)
        S.activeSeries=get(ui.seriesList,'Value');
        syncSeriesControls();updateEditor();
    end
    function removeSeries(~,~)
        if isempty(S.datasets),return;end
        ix=min(get(ui.seriesList,'Value'),numel(S.datasets));
        S.datasets(ix)=[];
        S.activeSeries=max(1,min(ix,numel(S.datasets)));
        syncSeriesControls();redraw();status('Dataset removed.');
    end
    function renameSeries(~,~)
        if isempty(S.datasets),return;end
        name=strtrim(get(ui.seriesName,'String'));
        if isempty(name),return;end
        S.datasets(S.activeSeries).name=name;
        syncSeriesControls();redraw();
    end
    function legendChanged(~,~)
        S.legendOn=logical(get(ui.legendCheck,'Value'));
        redraw();
    end

    function catalog=buildCatalog()
        tf=figure('Visible','off','HandleVisibility','off');
        cleanup=onCleanup(@()delete(tf)); %#ok<NASGU>
        ta=axes('Parent',tf);
        th=plot(ta,0,0,'o-');
        tl=legend(ta,th,'test');
        objs={th,ta,ta.XAxis,ta.YAxis,ta.XLabel,ta.YLabel,tl};
        cats={'Dataset','Axes','XRuler','YRuler','XLabel','YLabel','Legend'};
        catalog=struct('key',{},'category',{},'property',{},'editable',{});
        for cc=1:numel(cats)
            try
                writable=set(objs{cc});
                allProperties=get(objs{cc});
                props=sort(fieldnames(allProperties));
                for pp=1:numel(props)
                    prop=props{pp};
                    % Include read-only properties for visibility, but only
                    % permit editing settable scalar/text/numeric values.
                    val=allProperties.(prop);
                    simple=ischar(val)||isstring(val)||isnumeric(val)|| ...
                        islogical(val)||isa(val,'matlab.lang.OnOffSwitchState');
                    if isnumeric(val) && numel(val)>100,simple=false;end
                    allowed=isfield(writable,prop) && simple;
                    catalog(end+1)=entry(cats{cc},prop,allowed); %#ok<AGROW>
                end
            catch
            end
        end
        for p={'WidthMM','HeightMM','LeftMM','RightMM','BottomMM','TopMM'}
            catalog(end+1)=entry('Panel',p{1},true); %#ok<AGROW>
        end
        for p={'XTickGapPt','YTickGapPt'}
            catalog(end+1)=entry('Custom',p{1},true); %#ok<AGROW>
        end
        for category={'XTicks','YTicks'}
            for field={'Mode','Positions','Format','Decimals','Prefix', ...
                    'Suffix','CustomFormat','Labels'}
                catalog(end+1)=entry(category{1},field{1},true); %#ok<AGROW>
            end
        end
    end

    function e=entry(cat,prop,editable)
        e=struct('key',[cat '|' prop], 'category',cat, ...
            'property',prop,'editable',editable);
    end

    function filterProperties(~,~)
        groupNames=get(ui.group,'String');
        category=groupNames{get(ui.group,'Value')};
        filter=lower(strtrim(get(ui.search,'String')));
        found=strcmp({S.catalog.category},category);
        keys={S.catalog.key};
        if ~isempty(filter)
            found=found & cellfun(@(s)contains(lower(s),filter),keys);
        end
        c=S.catalog(found);
        S.displayCatalog={c.key};
        labels=cell(numel(c),1);
        for ii=1:numel(c)
            if c(ii).editable
                labels{ii}=c(ii).property;
            else
                labels{ii}=[c(ii).property '   [complex / view only]'];
            end
        end
        if isempty(labels),labels={'(No matching properties)'};end
        set(ui.properties,'String',labels,'Value',1);
    end

    function addWorkspace(~,~)
        if isempty(S.displayCatalog),return;end
        ix=get(ui.properties,'Value');
        key=S.displayCatalog{ix};
        if ~ismember(key,S.workspace)
            S.workspace{end+1}=key;
        end
        ui.tabs.SelectedTab=ui.tabEditor;
        updateEditor();
        status(['Added control: ' key]);
    end

    function updateEditor()
        data=cell(numel(S.workspace),2);
        for kk=1:numel(S.workspace)
            key=S.workspace{kk};data{kk,1}=strrep(key,'|',' / ');
            try data{kk,2}=formatValue(getProperty(key));
            catch,data{kk,2}='(unavailable)';end
        end
        set(ui.editor,'Data',data);
        if isempty(S.editorSelection)||S.editorSelection(1)>numel(S.workspace)
            S.editorSelection=[];
        end
        showEditorInput();
    end

    function editorCellSelected(~,event)
        if isempty(event.Indices),return;end
        S.editorSelection=event.Indices(1);
        showEditorInput();
    end

    function yes=isColorKey(key)
        parts=strsplit(key,'|');prop=lower(parts{2});
        yes=strcmp(prop,'color') || endsWith(prop,'color');
    end

    function [opts,labels]=propertyChoices(key)
        % Values are the exact MATLAB property codes; labels show a visual
        % preview and human-readable name. Saved figures use raw codes.
        opts={};labels={};parts=strsplit(key,'|');
        group=parts{1};prop=parts{2};
        if isColorKey(key),return;end
        % MATLAB stores many graphics on/off properties as an enum rather
        % than a text scalar. Always offer a clear dropdown for these.
        try
            current=getProperty(key);
            if isa(current,'matlab.lang.OnOffSwitchState') || ...
                    (islogical(current) && isscalar(current))
                opts={'on','off'};labels={'On  (enabled)','Off  (disabled)'};
                return;
            end
        catch
        end
        switch key
            case 'Dataset|Marker'
                opts={'none','o','+','*','.','x','s','d','^','v','>','<','p','h','_','|'};
                labels={'None (no marker)','○  Circle','+  Plus', ...
                    '✳  Asterisk','•  Point','×  Cross','□  Square', ...
                    '◇  Diamond','△  Triangle up','▽  Triangle down', ...
                    '▷  Triangle right','◁  Triangle left', ...
                    '☆  Pentagram','⬡  Hexagram','━  Horizontal bar', ...
                    '│  Vertical bar'};
            case 'Dataset|LineStyle'
                opts={'-','--',':','-.','none'};
                labels={'━━━━━  Solid','━ ━ ━  Dashed', ...
                    '··········  Dotted','━ · ━ ·  Dash-dot', ...
                    'None (no connecting line)'};
            case {'XTicks|Mode','YTicks|Mode'}
                opts={'auto','manual'};
            case {'XTicks|Format','YTicks|Format'}
                opts={'auto','number','general','scientific','engineering','si','custom'};
        end
        if ~isempty(opts)
            if isempty(labels),labels=opts;end
            return;
        end
        try
            if strcmp(group,'Dataset')
                if isempty(S.datasets),return;end
                target=S.plotHandles{S.activeSeries};
            elseif any(strcmp(group,{'Panel','Custom','XTicks','YTicks'}))
                return;
            else
                target=getTarget(group,ui.ax);
            end
            if isempty(target)||~isgraphics(target),return;end
            choices=set(target,prop);
            if isstring(choices),choices=cellstr(choices);end
            if ischar(choices) && size(choices,1)>1
                choices=cellstr(choices);
            end
            if iscell(choices) && numel(choices)>=2 && numel(choices)<=50
                try
                    opts=cellfun(@normalizeEnumChoice,choices(:)', ...
                        'UniformOutput',false);
                    labels=opts;
                    if isequal(sort(opts),sort({'on','off'}))
                        opts={'on','off'};labels={'On  (enabled)','Off  (disabled)'};
                    end
                catch
                    opts={};labels={};
                end
            end
        catch
            opts={};labels={};
        end
    end

    function value=normalizeEnumChoice(value)
        if isa(value,'matlab.lang.OnOffSwitchState')
            if logical(value),value='on';else,value='off';end
        elseif isstring(value)
            value=char(value);
        end
        if ~ischar(value),error('Unsupported enumeration value');end
        value=regexprep(value,'^matlab\.lang\.OnOffSwitchState\.','');
    end

    function populateColorChoices(key,currentValue)
        choices={};labels={};
        isMarkerColor=any(strcmp(key, ...
            {'Dataset|MarkerFaceColor','Dataset|MarkerEdgeColor'}));
        if isMarkerColor
            choices{end+1}='none'; %#ok<AGROW>
            labels{end+1}='None  (no fill / no outline)'; %#ok<AGROW>
            choices{end+1}='auto'; %#ok<AGROW>
            labels{end+1}='Auto  (MATLAB default)'; %#ok<AGROW>
        end
        % Keep an explicit RGB that is absent from the current palette.
        matched=false;
        if isnumeric(currentValue) && numel(currentValue)==3
            for j=1:numel(S.palette.colors)
                if all(abs(double(currentValue(:)')-S.palette.colors(j).rgb)<1e-7)
                    matched=true;break;
                end
            end
            if ~matched
                choices{end+1}=double(currentValue(:)'); %#ok<AGROW>
                labels{end+1}=['Current custom  ' rgbHex(currentValue)]; %#ok<AGROW>
            end
        end
        for j=1:numel(S.palette.colors)
            c=S.palette.colors(j);
            choices{end+1}=j; %#ok<AGROW>
            labels{end+1}=sprintf('■ %02d  %s  (%s)', ...
                j,c.name,rgbHex(c.rgb)); %#ok<AGROW>
        end
        if isempty(labels)
            labels={'(No palette colors)'};choices={[]};
        end
        chosen=1;
        for j=1:numel(choices)
            candidate=choices{j};
            if ischar(candidate) && (ischar(currentValue)||isstring(currentValue)) ...
                    && strcmpi(candidate,char(currentValue))
                chosen=j;break;
            elseif isnumeric(candidate) && isnumeric(currentValue) ...
                    && numel(candidate)==3 && numel(currentValue)==3 ...
                    && all(abs(double(candidate(:)')-double(currentValue(:)'))<1e-7)
                chosen=j;break;
            elseif isnumeric(candidate) && isscalar(candidate) ...
                    && isnumeric(currentValue) && numel(currentValue)==3 ...
                    && all(abs(S.palette.colors(candidate).rgb-double(currentValue(:)'))<1e-7)
                chosen=j;break;
            end
        end
        ui.quickPaletteValues=choices;
        set(ui.quickPalette,'String',labels,'Value',chosen);
    end

    function showEditorInput()
        set(ui.valueEdit,'Visible','off');
        set(ui.valueChoice,'Visible','off');
        set(ui.quickPalette,'Visible','off');
        set(ui.openPaletteButton,'Enable','off');
        ui.editorInputMode='none';
        if isempty(S.editorSelection)||S.editorSelection(1)>numel(S.workspace)
            set(ui.editorHelp,'String','Select a property row to edit.');
            return;
        end
        key=S.workspace{S.editorSelection(1)};
        try value=getProperty(key);catch,value='';end
        if isColorKey(key)
            ui.editorInputMode='palette';
            populateColorChoices(key,value);
            set(ui.quickPalette,'Visible','on');
            set(ui.openPaletteButton,'Enable','on');
            set(ui.editorHelp,'String', ...
                'Choose None/Auto (for markers) or a fixed RGB palette swatch.');
        else
            [choices,choiceLabels]=propertyChoices(key);
            if ~isempty(choices)
                ui.editorInputMode='choice';
                ui.valueChoiceValues=choices;
                set(ui.valueChoice,'String',choiceLabels,'Visible','on');
                ix=find(strcmpi(choices,formatValue(value)),1);
                if isempty(ix),ix=1;end
                set(ui.valueChoice,'Value',ix);
                set(ui.editorHelp,'String','Select a style preview, then Apply.');
            else
                ui.editorInputMode='edit';
                set(ui.valueEdit,'String',formatValue(value),'Visible','on');
                set(ui.editorHelp,'String','Enter a value, then Apply.');
            end
        end
    end

    function applyValue(~,~)
        if isempty(S.editorSelection),return;end
        ix=S.editorSelection(1);
        if ix>numel(S.workspace),return;end
        try
            switch ui.editorInputMode
                case 'palette'
                    useQuickPalette();return;
                case 'choice'
                    value=ui.valueChoiceValues{get(ui.valueChoice,'Value')};
                case 'edit'
                    value=get(ui.valueEdit,'String');
                otherwise
                    return;
            end
            applyProperty(S.workspace{ix},value);
            redraw();syncAxisTypographyControls();
            status('Property updated.');
        catch err
            errordlg(err.message,'Invalid property value');updateEditor();
        end
    end

    function pickColor(~,~)
        key=editorSelectedColorKey();
        if isempty(key),return;end
        S.palettePendingKey=key;
        ui.tabs.SelectedTab=ui.tabPalette;
        status(['Choose a swatch to set ' strrep(key,'|',' / ') ' (fixed RGB).']);
    end

    function key=editorSelectedColorKey()
        key='';
        if isempty(S.editorSelection) || S.editorSelection(1)>numel(S.workspace)
            warndlg('Select a color property in My Editor first.');return;
        end
        candidate=S.workspace{S.editorSelection(1)};
        if ~isColorKey(candidate)
            warndlg(['Select a color property: Dataset / Color, ' ...
                'MarkerFaceColor, or MarkerEdgeColor.']);return;
        end
        key=candidate;
    end

    function quickPaletteChanged(~,~)
        ix=get(ui.quickPalette,'Value');
        if ix<1 || ix>numel(ui.quickPaletteValues),return;end
        choice=ui.quickPaletteValues{ix};
        if isnumeric(choice) && isscalar(choice) && ...
                choice>=1 && choice<=numel(S.palette.colors)
            S.paletteSelected=choice;
            S.palettePage=ceil(S.paletteSelected/12);
            syncPaletteControls();
        end
    end

    function useQuickPalette(~,~)
        key=editorSelectedColorKey();
        if isempty(key),return;end
        ix=get(ui.quickPalette,'Value');
        if ix<1 || ix>numel(ui.quickPaletteValues),return;end
        choice=ui.quickPaletteValues{ix};
        if isnumeric(choice) && isscalar(choice) && ...
                choice>=1 && choice<=numel(S.palette.colors)
            applyPaletteRGB(key,choice);
        elseif (ischar(choice) && any(strcmpi(choice,{'none','auto'}))) ...
                || (isnumeric(choice) && numel(choice)==3)
            try
                applyProperty(key,choice);
                if startsWith(key,'Dataset|') && ~isempty(S.datasets)
                    S.datasets(S.activeSeries).sourceColorName='';
                end
                redraw();
                status([strrep(key,'|',' / ') ' set to ' formatValue(choice) '.']);
            catch err
                errordlg(err.message,'Color assignment failed');
            end
        end
    end

    function applyPaletteRGB(key,ix)
        c=S.palette.colors(ix);
        try
            applyProperty(key,c.rgb);
            if startsWith(key,'Dataset|') && ~isempty(S.datasets)
                S.datasets(S.activeSeries).sourceColorName=c.name;
            end
            redraw();
            status(sprintf('%s <= %s (%s), copied as fixed RGB.', ...
                strrep(key,'|',' / '),c.name,rgbHex(c.rgb)));
        catch err
            errordlg(err.message,'Palette assignment failed');
        end
    end

    function removeWorkspace(~,~)
        if isempty(S.editorSelection),return;end
        ix=S.editorSelection(1);
        if ix>numel(S.workspace),return;end
        S.workspace(ix)=[];S.editorSelection=[];
        updateEditor();
    end

    function resetProperty(~,~)
        if isempty(S.editorSelection),return;end
        key=S.workspace{S.editorSelection(1)};
        split=strsplit(key,'|');group=split{1};prop=split{2};
        if strcmp(group,'Dataset')
            if isempty(S.datasets),return;end
            if isfield(S.styleDefaults,prop)
                S.datasets(S.activeSeries).style.(prop)=S.styleDefaults.(prop);
            else
                warndlg('The master has no explicit default for this dataset property.');return;
            end
        elseif any(strcmp(group,{'XTicks','YTicks'}))
            defaults=defaultTickSettings();
            S.ticks.(group(1)).(prop)=defaults.(prop);
            syncTickControls();
        elseif isfield(S.overrides,group)
            if isfield(S.overrides.(group),prop)
                S.overrides.(group)=rmfield(S.overrides.(group),prop);
            end
            if strcmp(group,'Axes') && any(strcmp(prop,{'XLim','YLim'}))
                S.overrides.Axes.([prop 'Mode'])='auto';
                syncLimitControls();
            end
        else
            warndlg('Panel/custom values have no separate reset value.');return;
        end
        redraw();status('Removed figure-specific override.');
    end

    function val=getProperty(key)
        parts=strsplit(key,'|');group=parts{1};prop=parts{2};
        if strcmp(group,'Dataset')
            if isempty(S.datasets),error('No selected dataset.');end
            idx=S.activeSeries;
            if isfield(S.datasets(idx).style,prop)
                val=S.datasets(idx).style.(prop);return;
            end
            if idx<=numel(S.plotHandles) && ~isempty(S.plotHandles{idx}) && isgraphics(S.plotHandles{idx})
                val=get(S.plotHandles{idx},prop);return;
            end
            error('Series property unavailable.');
        elseif strcmp(group,'Panel')
            field=[lower(prop(1)) prop(2:end)];
            val=S.panel.(field);return;
        elseif strcmp(group,'Custom')
            val=S.custom.(prop);return;
        elseif any(strcmp(group,{'XTicks','YTicks'}))
            side=group(1);
            val=S.ticks.(side).(prop);return;
        else
            if isfield(S.overrides.(group),prop)
                val=S.overrides.(group).(prop);return;
            end
            target=getTarget(group,ui.ax);
            val=get(target,prop);
        end
    end

    function target=getTarget(group,ax)
        switch group
            case 'Axes',target=ax;
            case 'XRuler',target=ax.XAxis;
            case 'YRuler',target=ax.YAxis;
            case 'XLabel',target=ax.XLabel;
            case 'YLabel',target=ax.YLabel;
            case 'Legend'
                target=legend(ax);
            otherwise,error('Unknown property group: %s',group);
        end
    end

    function applyProperty(key,newRaw)
        parts=strsplit(key,'|');group=parts{1};prop=parts{2};
        ix=find(strcmp({S.catalog.key},key),1);
        if isempty(ix) || ~S.catalog(ix).editable
            error('This complex property is view-only in the generic editor.');
        end
        if strcmp(group,'Dataset')
            if isempty(S.datasets),error('Import or select a dataset first.');end
            target=S.plotHandles{S.activeSeries};
        elseif any(strcmp(group,{'Panel','Custom','XTicks','YTicks'}))
            target=[];
        else
            target=getTarget(group,ui.ax);
        end
        if strcmp(group,'Panel')
            field=[lower(prop(1)) prop(2:end)];
            old=S.panel.(field);
        elseif strcmp(group,'Custom')
            old=S.custom.(prop);
        elseif any(strcmp(group,{'XTicks','YTicks'}))
            old=S.ticks.(group(1)).(prop);
        else
            old=get(target,prop);
        end
        if any(strcmp(group,{'XTicks','YTicks'}))
            setTickProperty(group(1),prop,newRaw);
            syncTickControls();
            return;
        end
        if strcmp(group,'Axes') && any(strcmp(prop,{'XTick','YTick'}))
            setTickProperty(prop(1),'Positions',newRaw);
            syncTickControls();
            return;
        end
        if strcmp(group,'Axes') && any(strcmp(prop,{'XTickMode','YTickMode'}))
            setTickProperty(prop(1),'Mode',newRaw);
            syncTickControls();
            return;
        end
        % RGB triplets are numeric, but marker colors also accept the
        % *textual* states 'none' and 'auto'. Do not send those states through
        % the numeric-vector parser when an RGB color was selected before.
        if strcmp(group,'Dataset') && ...
                any(strcmp(prop,{'MarkerFaceColor','MarkerEdgeColor'})) && ...
                (ischar(newRaw)||isstring(newRaw)) && ...
                any(strcmpi(strtrim(char(newRaw)),{'none','auto'}))
            val=lower(strtrim(char(newRaw)));
        else
            val=parseValue(newRaw,old);
        end
        if strcmp(group,'Panel')
            previous=S.panel;
            S.panel.(field)=val;
            try
                validatePanel();
            catch err
                S.panel=previous;
                rethrow(err);
            end
            S.panel.preset='custom';syncPanelControls();
        elseif strcmp(group,'Custom')
            if ~isscalar(val) || ~isfinite(val),error('Enter a finite number.');end
            S.custom.(prop)=val;
        elseif strcmp(group,'Dataset')
            updated=S.datasets(S.activeSeries).style;
            updated.(prop)=val;
            if strcmp(prop,'MarkerEdgeColor')
                % Some MATLAB releases reject native edge color 'none'.
                % Validate either the native state or its visual fallback.
                applyMarkerEdgeColor(target,val,updated);
            else
                set(target,prop,val); % Validate against MATLAB first.
                if strcmp(prop,'MarkerFaceColor') && ...
                        isfield(updated,'MarkerEdgeColor') && ...
                        (ischar(updated.MarkerEdgeColor)||isstring(updated.MarkerEdgeColor)) && ...
                        strcmpi(char(updated.MarkerEdgeColor),'none')
                    % Recalculate the edge fallback if the fill changes.
                    applyMarkerEdgeColor(target,'none',updated);
                end
            end
            S.datasets(S.activeSeries).style=updated;
        else
            set(target,prop,val);
            S.overrides.(group).(prop)=val;
            if strcmp(group,'Axes') && any(strcmp(prop,{'XLim','YLim'}))
                S.overrides.Axes.([prop 'Mode'])='manual';
                syncLimitControls();
            elseif strcmp(group,'Axes') && any(strcmp(prop,{'XLimMode','YLimMode'}))
                if strcmpi(val,'auto')
                    oldProp=prop(1:end-4);
                    if isfield(S.overrides.Axes,oldProp)
                        S.overrides.Axes=rmfield(S.overrides.Axes,oldProp);
                    end
                end
                syncLimitControls();
            end
        end
    end

    function setTickProperty(side,prop,raw)
        t=S.ticks.(side);
        switch prop
            case 'Positions'
                t.Positions=parseTickSequence(raw);
                t.Mode=onOffTick(~isempty(t.Positions));
            case 'Mode'
                if isnumeric(raw)
                    error('Tick mode must be auto or manual.');
                end
                raw=lower(strtrim(char(raw)));
                if ~any(strcmp(raw,{'auto','manual'}))
                    error('Tick mode must be auto or manual.');
                end
                if strcmp(raw,'manual') && isempty(t.Positions)
                    error('First enter tick positions, e.g. 0:0.5:3.');
                end
                t.Mode=raw;
            case 'Format'
                val=lower(strtrim(char(raw)));
                if ~any(strcmp(val,{'auto','number','general', ...
                        'scientific','engineering','si','custom'}))
                    error('Choose auto, number, general, scientific, engineering, si or custom.');
                end
                t.Format=val;
            case 'Decimals'
                val=str2double(num2str(raw));
                if ~isfinite(val)||fix(val)~=val||val<0||val>12
                    error('Decimals must be an integer from 0 to 12.');
                end
                t.Decimals=val;
            case 'CustomFormat'
                validateFormat(raw);
                t.CustomFormat=char(raw);
            case {'Prefix','Suffix','Labels'}
                t.(prop)=char(raw);
            otherwise,error('Unknown tick property.');
        end
        if ~isempty(t.Labels) && strcmpi(t.Mode,'manual')
            labels=strsplit(t.Labels,'|');
            if numel(labels)~=numel(t.Positions)
                error('Custom labels must match the number of tick positions.');
            end
        end
        S.ticks.(side)=t;
    end

    function val=onOffTick(manual)
        if manual,val='manual';else,val='auto';end
    end

    function val=parseValue(raw,old)
        if isnumeric(raw)||islogical(raw),val=raw;return;end
        if isstring(raw),raw=char(raw);end
        raw=strtrim(raw);
        raw=regexprep(raw,'^matlab\.lang\.OnOffSwitchState\.','');
        if isa(old,'matlab.lang.OnOffSwitchState')
            if any(strcmpi(raw,{'on','1','true','yes'})),val='on';
            elseif any(strcmpi(raw,{'off','0','false','no'})),val='off';
            else,error('Choose On or Off.');end
        elseif islogical(old)
            if any(strcmpi(raw,{'1','true','on','yes'})),val=true;
            elseif any(strcmpi(raw,{'0','false','off','no'})),val=false;
            else,error('Enter true/false or on/off.');end
        elseif isnumeric(old)
            clean=regexprep(raw,'[\[\],;]',' ');
            tokens=regexp(strtrim(clean),'\s+','split');
            if numel(tokens)==1 && isempty(tokens{1}),error('Numeric value required.');end
            nums=cellfun(@str2double,tokens);
            if any(isnan(nums)),error('Expected a number or numeric vector, e.g. [0.2 0.5 0.8].');end
            if isscalar(old) && numel(nums)~=1
                error('This property needs a scalar number.');
            end
            if numel(old)>1 && numel(nums)~=numel(old)
                error('This property needs %d numeric values.',numel(old));
            end
            val=reshape(nums,size(old));
        elseif ischar(old)||isstring(old)
            val=raw;
        else
            error('This property requires a complex MATLAB value and is view-only.');
        end
    end

    function str=formatValue(val)
        if isa(val,'matlab.lang.OnOffSwitchState')
            if logical(val),str='on';else,str='off';end
        elseif ischar(val),str=val;
        elseif isstring(val),str=char(val);
        elseif isnumeric(val)||islogical(val)
            if isempty(val),str='[]';
            elseif numel(val)>12,str=sprintf('[%s numeric]',mat2str(size(val)));
            else,str=mat2str(val,4);
            end
        else
            str=['<' class(val) '>'];
        end
    end

    function loadPalette(~,~)
        [f,p]=uigetfile({'*.xml','CorelDRAW palette (*.xml)'},'Open palette');
        if isequal(f,0),return;end
        try
            colors=readCorelPalette(fullfile(p,f));
            if isempty(colors),error('No compatible RGB/sRGB color entries were found.');end
            S.palette=struct('path',fullfile(p,f),'name',f,'colors',colors);
            S.paletteSelected=1;S.palettePage=1;S.palettePendingKey='';
            syncPaletteControls();
            updateEditor();
            status(sprintf('Loaded %d RGB colors. Existing series colors unchanged.',numel(colors)));
        catch err
            errordlg(err.message,'Palette import failed');
        end
    end

    function colors=readCorelPalette(filepath)
        dom=xmlread(filepath);
        nodes=dom.getElementsByTagName('color');
        colors=struct('name',{},'rgb',{});
        for ii=0:nodes.getLength()-1
            node=nodes.item(ii);
            cs=lower(strtrim(char(node.getAttribute('cs'))));
            tints=char(node.getAttribute('tints'));
            if isempty(tints),continue;end
            components=regexp(tints,'\s*[,; ]\s*','split');
            rgb=cellfun(@str2double,components);
            if any(~isfinite(rgb)),continue;end
            if any(strcmp(cs,{'rgb','srgb'})) && numel(rgb)>=3
                rgb=rgb(1:3);
                if any(rgb>1),rgb=rgb/255;end
            else
                % CMYK, Lab and spot colors require calibrated transforms.
                % Skip rather than silently substitute inaccurate colors.
                continue;
            end
            if any(rgb<0)||any(rgb>1),continue;end
            name=strtrim(char(node.getAttribute('name')));
            if isempty(name),name=sprintf('Swatch %d',ii+1);end
            colors(end+1)=struct('name',name,'rgb',rgb(:)'); %#ok<AGROW>
        end
    end

    function prevPalettePage(~,~)
        S.palettePage=max(1,S.palettePage-1);syncPaletteControls();
    end
    function nextPalettePage(~,~)
        S.palettePage=min(ceil(numel(S.palette.colors)/12),S.palettePage+1);
        syncPaletteControls();
    end
    function swatchClicked(src,~)
        ix=get(src,'UserData');
        if isempty(ix)||ix>numel(S.palette.colors),return;end
        S.paletteSelected=ix;
        % The editor palette popup contains extra None/Auto entries, so
        % its index is not the same as the visual palette swatch index.
        syncPaletteControls();
        if ~isempty(S.palettePendingKey)
            key=S.palettePendingKey;
            S.palettePendingKey='';
            applyPaletteRGB(key,ix);
            ui.tabs.SelectedTab=ui.tabEditor;
        end
    end

    function assignColor(~,~)
        if isempty(S.datasets),warndlg('Select or import a dataset first.');return;end
        if isempty(S.palette.colors),return;end
        modes={'Color','MarkerFaceColor','MarkerEdgeColor'};
        key=['Dataset|' modes{get(ui.colorTarget,'Value')}];
        applyPaletteRGB(key,S.paletteSelected);
    end

    function panelPresetChanged(~,~)
        switch get(ui.panelPreset,'Value')
            case 1
                S.panel.preset='normal';S.panel.widthMM=85;S.panel.heightMM=65;
            case 2
                S.panel.preset='double';S.panel.widthMM=175;S.panel.heightMM=75;
            otherwise
                S.panel.preset='custom';
        end
        syncPanelControls();redraw();
    end

    function panelEdited(~,~)
        prev=S.panel;
        for kk=1:numel(fields)
            item=fields{kk};
            v=str2double(get(ui.panelEdits.(item),'String'));
            if isnan(v),errordlg('Panel dimensions must be numeric.');S.panel=prev;syncPanelControls();return;end
            S.panel.(item)=v;
        end
        try validatePanel();catch err
            S.panel=prev;syncPanelControls();errordlg(err.message);return;
        end
        S.panel.preset='custom';syncPanelControls();redraw();
    end

    function validatePanel()
        p=S.panel;
        v=[p.widthMM,p.heightMM,p.leftMM,p.rightMM,p.bottomMM,p.topMM];
        if any(~isfinite(v)) || p.widthMM<=0 || p.heightMM<=0 || any(v(3:end)<0)
            error('Panel dimensions must be positive and margins nonnegative.');
        end
        if p.widthMM<5 || p.heightMM<5
            error('Plot area must be at least 5 mm in both dimensions.');
        end
    end

    function saveRecipe(~,~)
        masterNameEdited([],[]);masterNotesEdited([],[]);
        exportSettingsEdited([],[]);
        try captureVisibleFontSizes();catch err
            errordlg(err.message,'Cannot save figure');return;
        end
        [f,p]=uiputfile('*.json','Save reproducible figure recipe', ...
            fullfile(pwd,'Figure_1a.json'));
        if isequal(f,0),return;end
        [~,base]=fileparts(f);
        jsonPath=fullfile(p,[base '.json']);
        cacheName=[base '_data.mat'];
        cachePath=fullfile(p,cacheName);
        cachedDatasets=S.datasets; %#ok<NASGU>
        save(cachePath,'cachedDatasets','-v7');
        recipe=struct('format','PublicationPlotEditorRecipe','version',7, ...
            'masterName',S.masterName,'notes',S.masterNotes,'workbook',S.workbook, ...
            'panel',S.panel,'export',S.export, ...
            'overrides',S.overrides,'custom',S.custom, ...
            'ticks',S.ticks, ...
            'styleDefaults',S.styleDefaults,'legendOn',S.legendOn, ...
            'workspace',{S.workspace},'dataSnapshot',cacheName, ...
            'paletteName',S.palette.name,'paletteSource',S.palette.path, ...
            'paletteSnapshot',S.palette.colors);
        series=S.datasets;
        for kk=1:numel(series)
            series(kk).x=[];series(kk).y=[];
        end
        recipe.datasets=series;
        writeJSON(jsonPath,recipe);
        S.recipePath=jsonPath;
        status(sprintf('Saved full figure and data snapshot: %s',jsonPath));
    end

    function loadRecipe(~,~)
        [f,p]=uigetfile('*.json','Load figure recipe');
        if isequal(f,0),return;end
        file=fullfile(p,f);
        try
            obj=jsondecode(fileread(file));
            if ~isfield(obj,'format')||~strcmp(obj.format,'PublicationPlotEditorRecipe')
                error('This is not a Publication Plot Editor figure recipe.');
            end
            datasets=obj.datasets;
            if isempty(datasets)
                datasets=S.datasets([]);
            end
            if isfield(obj,'dataSnapshot') && ~isempty(obj.dataSnapshot)
                cachePath=fullfile(p,obj.dataSnapshot);
                if exist(cachePath,'file')
                    temp=load(cachePath,'cachedDatasets');
                    if isfield(temp,'cachedDatasets') && numel(temp.cachedDatasets)==numel(datasets)
                        for kk=1:numel(datasets)
                            datasets(kk).x=temp.cachedDatasets(kk).x;
                            datasets(kk).y=temp.cachedDatasets(kk).y;
                        end
                    end
                end
            end
            S.masterName=char(obj.masterName);
            if isfield(obj,'notes')
                S.masterNotes=notesAsText(obj.notes);
            else
                S.masterNotes='';
            end
            S.workbook=char(obj.workbook);
            S.panel=normalizePanel(obj.panel,obj.version);
            if isfield(obj,'export')
                S.export=normalizeExportSettings(obj.export);
            else
                S.export=defaultExportSettings();
            end
            S.overrides=obj.overrides;
            S.custom=normalizeCustomSettings(obj.custom);
            if isfield(obj,'ticks')
                S.ticks=normalizeTickSettings(obj.ticks);
            else
                S.ticks=struct('X',defaultTickSettings(), ...
                    'Y',defaultTickSettings());
            end
            S.styleDefaults=obj.styleDefaults;
            S.legendOn=logical(obj.legendOn);
            S.workspace=normalizeWorkspace(obj.workspace);
            S.workspace=compactLegacyWorkspace(S.workspace,obj);
            S.workspace=S.workspace(ismember(S.workspace,{S.catalog.key}));
            if isfield(obj,'paletteSnapshot') && ~isempty(obj.paletteSnapshot)
                S.palette.colors=obj.paletteSnapshot;
                S.palette.name=char(obj.paletteName);
                S.palette.path=char(obj.paletteSource);
                S.paletteSelected=1;S.palettePage=1;
            end
            S.datasets=datasets;
            S.activeSeries=1;
            S.recipePath=file;
            updateEverything();
            status('Figure reopened: data and formatting restored.');
        catch err
            errordlg(err.message,'Recipe load failed');
        end
    end

    function saveTemplate(~,~)
        masterNameEdited([],[]);masterNotesEdited([],[]);
        exportSettingsEdited([],[]);
        try captureVisibleFontSizes();catch err
            errordlg(err.message,'Cannot save master');return;
        end
        [f,p]=uiputfile('*.json','Save master style (no data)', ...
            fullfile(pwd,'StandardScientific.json'));
        if isequal(f,0),return;end
        % A master should capture the selected curve's actual design, not
        % only the original built-in line defaults.
        masterStyle=S.styleDefaults;
        if ~isempty(S.datasets)
            masterStyle=S.datasets(S.activeSeries).style;
        end
        masterTicks=S.ticks;
        for axisSide={'X','Y'}
            k=axisSide{1};masterTicks.(k).Mode='auto';
            masterTicks.(k).Positions=[];masterTicks.(k).Labels='';
            masterTicks.(k).Prefix='';masterTicks.(k).Suffix='';
        end
        obj=struct('format','PublicationPlotEditorMaster','version',7, ...
            'name',S.masterName,'notes',S.masterNotes,'panel',S.panel, ...
            'export',S.export, ...
            'overrides',styleOnlyOverrides(S.overrides),'custom',S.custom, ...
            'ticks',masterTicks, ...
            'styleDefaults',masterStyle,'legendOn',S.legendOn, ...
            'workspace',{S.workspace},'paletteSource',S.palette.path, ...
            'paletteSnapshot',S.palette.colors);
        writeJSON(fullfile(p,f),obj);
        S.styleDefaults=masterStyle;
        status('Master style saved (selected series appearance, no data).');
    end

    function loadTemplate(~,~)
        [f,p]=uigetfile('*.json','Apply master style (no data)');
        if isequal(f,0),return;end
        try
            obj=jsondecode(fileread(fullfile(p,f)));
            if ~isfield(obj,'format')||~strcmp(obj.format,'PublicationPlotEditorMaster')
                error('Not a Publication Plot Editor master template.');
            end
            S.masterName=char(obj.name);
            if isfield(obj,'notes')
                S.masterNotes=notesAsText(obj.notes);
            else
                S.masterNotes='';
            end
            S.panel=normalizePanel(obj.panel,obj.version);
            if isfield(obj,'export')
                S.export=normalizeExportSettings(obj.export);
            else
                S.export=defaultExportSettings();
            end
            S.overrides=mergeStyleOverrides(S.overrides, ...
                styleOnlyOverrides(obj.overrides));
            S.custom=normalizeCustomSettings(obj.custom);
            if isfield(obj,'ticks')
                imported=normalizeTickSettings(obj.ticks);
                for axisSide={'X','Y'}
                    k=axisSide{1};
                    S.ticks.(k).Format=imported.(k).Format;
                    S.ticks.(k).Decimals=imported.(k).Decimals;
                    S.ticks.(k).CustomFormat=imported.(k).CustomFormat;
                end
            else
                S.ticks=struct('X',defaultTickSettings(), ...
                    'Y',defaultTickSettings());
            end
            S.styleDefaults=obj.styleDefaults;
            S.legendOn=logical(obj.legendOn);
            S.workspace=normalizeWorkspace(obj.workspace);
            S.workspace=compactLegacyWorkspace(S.workspace,obj);
            S.workspace=S.workspace(ismember(S.workspace,{S.catalog.key}));
            if isfield(obj,'paletteSnapshot') && ~isempty(obj.paletteSnapshot)
                S.palette.colors=obj.paletteSnapshot;
                S.palette.name='Palette snapshot from master';
                if isfield(obj,'paletteSource'),S.palette.path=char(obj.paletteSource);end
            end
            if ~isempty(S.datasets)
                choice=questdlg(['Apply master dataset style to all currently imported ' ...
                    'datasets? Colors will be copied as fixed RGB values.'], ...
                    'Master template','Keep existing styles','Apply master', ...
                    'Keep existing styles');
                if strcmp(choice,'Apply master')
                    for kk=1:numel(S.datasets)
                        S.datasets(kk).style=styleForSeries(kk);
                    end
                end
            end
            updateEverything();status('Applied master style; no data imported or replaced.');
        catch err
            errordlg(err.message,'Template load failed');
        end
    end

    function filtered=styleOnlyOverrides(input)
        % A master contains appearance. Plot-specific labels, extents and
        % numeric tick locations always remain with the figure recipe.
        filtered=input;
        for group={'XLabel','YLabel'}
            g=group{1};
            if isfield(filtered,g) && isfield(filtered.(g),'String')
                filtered.(g)=rmfield(filtered.(g),'String');
            end
        end
        if isfield(filtered,'Axes')
            for prop={'XLim','YLim','XLimMode','YLimMode', ...
                    'XTick','YTick','XTickLabel','YTickLabel', ...
                    'XTickMode','YTickMode','XTickLabelMode','YTickLabelMode'}
                k=prop{1};
                if isfield(filtered.Axes,k)
                    filtered.Axes=rmfield(filtered.Axes,k);
                end
            end
        end
    end

    function result=mergeStyleOverrides(existing,master)
        result=existing;
        for groups=fieldnames(master)'
            group=groups{1};
            if ~isfield(result,group),result.(group)=struct();end
            props=fieldnames(master.(group));
            for i=1:numel(props)
                p=props{i};
                result.(group).(p)=master.(group).(p);
            end
        end
    end

    function result=compactLegacyWorkspace(workspace,obj)
        result=workspace;
        if ~isfield(obj,'version') || obj.version>2,return;end
        % Version 0.2 put these controls redundantly in My Editor. Their
        % editable values are now collected in Axes & ticks or Panel.
        duplicates={'XRuler|FontSize','YRuler|FontSize', ...
            'XLabel|String','XLabel|FontSize', ...
            'YLabel|String','YLabel|FontSize', ...
            'Axes|XLim','Axes|YLim', ...
            'XTicks|Positions','XTicks|Format', ...
            'YTicks|Positions','YTicks|Format', ...
            'Panel|WidthMM','Panel|HeightMM'};
        result=result(~ismember(result,duplicates));
    end

    function writeJSON(filepath,obj)
        [fid,msg]=fopen(filepath,'w','n','UTF-8');
        if fid<0,error('Cannot save JSON: %s',msg);end
        cleanup=onCleanup(@()fclose(fid)); %#ok<NASGU>
        fprintf(fid,'%s\n',jsonencode(obj));
    end

    function applyEMFFontCorrection(ax)
        % Independent export-only calibration for X and Y rulers / labels.
        % Reapply label sizes AFTER rulers: MATLAB's label font multiplier
        % otherwise silently modifies axis-label typography.
        sizes=zeros(1,2);sides={'X','Y'};
        for j=1:2
            k=sides{j};label=ax.([k 'Label']);
            if isfield(S.overrides.([k 'Label']),'FontSize')
                sizes(j)=double(S.overrides.([k 'Label']).FontSize);
            else
                set(label,'FontUnits','points');sizes(j)=double(label.FontSize);
            end
        end
        if isprop(ax,'FontSizeMode')
            try ax.FontSizeMode='manual';catch,end
        end
        for j=1:2
            key=[sides{j} 'Tick'];
            ruler=ax.([sides{j} 'Axis']);
            try
                ruler.FontSize=ruler.FontSize*S.export.FontFactors.(key);
            catch err
                error('Independent %s tick font correction failed: %s', ...
                    sides{j},err.message);
            end
        end
        for j=1:2
            k=sides{j};label=ax.([k 'Label']);
            set(label,'FontUnits','points', ...
                'FontSize',sizes(j)*S.export.FontFactors.([k 'Label']));
        end
    end

    function exportFigure(~,~)
        exportSettingsEdited([],[]);
        try captureVisibleFontSizes();catch err
            errordlg(err.message,'Cannot export figure');return;
        end
        if isempty(S.datasets),warndlg('Import at least one Excel series first.');return;end
        [f,p]=uiputfile('*.emf','Choose export basename',fullfile(pwd,'Figure_1a.emf'));
        if isequal(f,0),return;end
        [~,base]=fileparts(f);
        emf=fullfile(p,[base '.emf']);
        pdf=fullfile(p,[base '.pdf']);
        dims=panelCanvasMM();w=dims(1)/10;h=dims(2)/10;
        ef=figure('Visible','off','Color','w','InvertHardcopy','off', ...
            'Units','centimeters','Position',[3 3 w h], ...
            'PaperUnits','centimeters','PaperSize',[w h], ...
            'PaperPosition',[0 0 w h],'PaperPositionMode','manual');
        cleanup=onCleanup(@()delete(ef)); %#ok<NASGU>
        ea=axes('Parent',ef,'Units','centimeters','Position',axesCM());
        drawPlot(ea,true);
        errors={};
        try
            print(ef,pdf,'-dpdf','-painters');
        catch err
            errors{end+1}=['PDF: ' err.message]; %#ok<AGROW>
        end
        % MATLAB .fig and PDF store the user-requested point sizes.
        % Font calibration affects only the EMF sent to CorelDRAW.
        try
            savefig(ef,fullfile(p,[base '.fig']));
        catch err
            errors{end+1}=['FIG: ' err.message]; %#ok<AGROW>
        end
        if ispc
            try
                if S.export.ApplyEMFCorrection
                    applyEMFFontCorrection(ea);
                end
                setInnerPositionConstraint(ea);
                set(ea,'Units','centimeters','Position',axesCM());
                print(ef,emf,'-dmeta','-painters');
            catch err
                errors{end+1}=['EMF: ' err.message]; %#ok<AGROW>
            end
        else
            errors{end+1}='EMF requires Windows.';
        end
        if isempty(errors)
            status(sprintf('Exported EMF/PDF/FIG. Plot %.1f x %.1f mm; canvas %.1f x %.1f mm.', ...
                S.panel.widthMM,S.panel.heightMM,dims(1),dims(2)));
        else
            warndlg(strjoin(errors,newline),'Some export formats failed');
            status('Export completed with warnings; check output files.');
        end
    end
end

% ---------------------- Standalone local helpers ----------------------

function result=notesAsText(value)
    % Handles the char matrix/cell array returned by multiline uicontrols.
    if iscell(value)
        result=strjoin(cellfun(@char,value,'UniformOutput',false),newline);
    elseif isstring(value)
        result=strjoin(cellstr(value),newline);
    elseif ischar(value)
        if size(value,1)>1
            result=strjoin(cellstr(value),newline);
        else
            result=value;
        end
    else
        result='';
    end
end

function result=normalizeWorkspace(input)
    if iscell(input)
        result=cellfun(@char,input,'UniformOutput',false);
    elseif isstring(input)
        result=cellstr(input);
    elseif ischar(input)
        result={input};
    else
        result={};
    end
end

function style=defaultStyle()
    style=struct('LineStyle','-','LineWidth',1.0,'Color',[0.15 0.4 0.7], ...
        'Marker','o','MarkerSize',5, ...
        'MarkerFaceColor','none','MarkerEdgeColor','auto', ...
        'Visible','on');
end

function colors=defaultPalette()
    rgb=[0.15 .39 .70; .83 .39 .19; .16 .55 .39; .53 .34 .68; ...
         .72 .25 .31; .14 .61 .68; .65 .51 .15; .2 .2 .2];
    names={'Blue','Orange','Green','Purple','Red','Teal','Ochre','Black'};
    colors=struct('name',{},'rgb',{});
    for i=1:size(rgb,1)
        colors(end+1)=struct('name',names{i},'rgb',rgb(i,:)); %#ok<AGROW>
    end
end

function textHandle=makeText(parent,str,rect)
    textHandle=uicontrol(parent,'Style','text','String',str, ...
        'Units','normalized','Position',rect, ...
        'HorizontalAlignment','left');
end

function h=makeButton(parent,label,pos,callback)
    h=uicontrol(parent,'Style','pushbutton','String',label, ...
        'Units','normalized','Position',pos,'Callback',callback);
end

function s=rgbHex(rgb)
    n=round(255*max(0,min(1,double(rgb(:)'))));
    s=sprintf('#%02X%02X%02X',n(1),n(2),n(3));
end

function panel=normalizePanel(panel,sourceVersion)
    % In v0.4 and earlier, width/height described the WHOLE canvas.
    % Convert once, retaining the same physical rectangle for the plot.
    if ~isfield(panel,'dimensionMode') || ...
            ~strcmpi(char(panel.dimensionMode),'plotArea')
        if nargin<2 || isempty(sourceVersion)
            sourceVersion=4;
        end
        if sourceVersion<=4
            panel.widthMM=panel.widthMM-panel.leftMM-panel.rightMM;
            panel.heightMM=panel.heightMM-panel.bottomMM-panel.topMM;
            if panel.widthMM<=0 || panel.heightMM<=0
                error('Legacy panel margins exceed the saved canvas dimensions.');
            end
        end
    end
    panel.dimensionMode='plotArea';
end

function result=defaultCustomSettings()
    % 'primary' = bottom/left tick marks with four border edges.
    % 'all' = native MATLAB Box='on' with four-sided ticks.
    result=struct('XTickGapPt',0,'YTickGapPt',0, ...
        'BoxTickSides','primary');
end

function result=normalizeCustomSettings(input)
    result=defaultCustomSettings();
    names=fieldnames(result);
    for k=1:numel(names)
        if isfield(input,names{k}),result.(names{k})=input.(names{k});end
    end
    if ~any(strcmp(result.BoxTickSides,{'primary','all'}))
        result.BoxTickSides='primary';
    end
end

function result=defaultExportSettings()
    % Initial empirical CorelDRAW 19 corrections. The values can be
    % recalibrated independently through Panel > Calibrate X/Y fonts.
    tick=7/7.92;
    label=(8/9.681)*(7/6.16);
    result=struct('ApplyEMFCorrection',true, ...
        'FontFactors',struct('XTick',tick,'YTick',tick, ...
        'XLabel',label,'YLabel',label));
end

function result=normalizeExportSettings(input)
    result=defaultExportSettings();
    if ~isstruct(input),return;end
    if isfield(input,'ApplyEMFCorrection')
        result.ApplyEMFCorrection=logical(input.ApplyEMFCorrection);
    end
    % Import native v0.7 per-object factors, or migrate the two global
    % v0.5/v0.6 correction factors to both corresponding axes.
    if isfield(input,'TickFontFactor')
        result.FontFactors.XTick=double(input.TickFontFactor);
        result.FontFactors.YTick=double(input.TickFontFactor);
    end
    if isfield(input,'AxisLabelFontFactor')
        v=double(input.AxisLabelFontFactor);
        if abs(v-8/9.681)<1e-6
            v=(8/9.681)*(7/6.16);
        end
        result.FontFactors.XLabel=v;
        result.FontFactors.YLabel=v;
    end
    if isfield(input,'FontFactors') && isstruct(input.FontFactors)
        keys=fieldnames(result.FontFactors);
        for j=1:numel(keys)
            k=keys{j};
            if isfield(input.FontFactors,k)
                result.FontFactors.(k)=double(input.FontFactors.(k));
            end
        end
    end
    keys=fieldnames(result.FontFactors);
    for j=1:numel(keys)
        v=result.FontFactors.(keys{j});
        if ~isnumeric(v)||~isscalar(v)||~isfinite(v)||v<=0||v>10
            error('Saved EMF font correction %s is invalid.',keys{j});
        end
    end
end

function result=defaultTickSettings()
    result=struct('Mode','auto','Positions',[], ...
        'Format','auto','Decimals',2,'Prefix','','Suffix','', ...
        'CustomFormat','%.2f','Labels','');
end

function ticks=normalizeTickSettings(input)
    ticks=struct('X',defaultTickSettings(),'Y',defaultTickSettings());
    for side={'X','Y'}
        k=side{1};
        if isfield(input,k)
            fields=fieldnames(ticks.(k));
            for j=1:numel(fields)
                field=fields{j};
                if isfield(input.(k),field)
                    ticks.(k).(field)=input.(k).(field);
                end
            end
        end
    end
end
