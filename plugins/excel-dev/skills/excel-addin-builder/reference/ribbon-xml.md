# Ribbon XML (customUI14) and its callbacks

The ribbon of an add-in is an XML file inside the workbook's zip,
`customUI/customUI14.xml`. Excel reads it when the workbook loads, builds the
controls, and calls VBA procedures **by name** when the user clicks. Those
procedures are called *callbacks*. Nothing checks that the names match the
code, and that is where most add-in bugs live.

Working example: `templates/customUI14-template.xml` together with
`templates/modRibbon.bas`.

---

## 1. Skeleton

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_OnLoad">
  <ribbon>
    <tabs>
      <tab id="tabMyAddIn" label="My Add-in">
        <group id="grpTools" label="Tools">
          <button id="btnRun" label="Run It" size="large" imageMso="TableInsert"
                  onAction="Ribbon_RunIt" screentip="Run It" supertip="What it does, in one sentence." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

- **Namespace.** `.../2009/07/customui` belongs to `customUI14.xml` (Office
  2010 and later). The older `customUI.xml` part uses `.../2006/01/customui`
  and cannot hold 2010-only features. Use customUI14 for anything new.
- **Built-in tab instead of your own:** `<tab idMso="TabHome">` adds your
  group to Home. Use `idMso` only for built-in controls and `id` for your own.
- `id` values must be **unique across the whole file**, separators included.
  A naming scheme helps: `grp*`, `btn*`, `mnu*`, `sep*`.

## 2. Controls you will actually use

| Element | Notes |
|---|---|
| `group` | A labelled block on the tab. A column holds three `size="normal"` controls. |
| `button` | `size="large"` or `"normal"`. `onAction` takes 1 parameter. |
| `separator` | A vertical line between controls in a group. Needs an `id`. |
| `toggleButton` | Has a pressed state. `onAction` gets **2** parameters (the new state). Use `getPressed` to show the current state. |
| `menu` | A dropdown list of buttons. `itemSize="large"` gives bigger items with descriptions. |
| `menuSeparator` | A line inside a menu. Optional `title="Section"` makes it a heading. |
| `splitButton` | Contains exactly one `button` (or `toggleButton`) plus one `menu`. The top half runs the button and the arrow opens the menu. |
| `dropDown` / `comboBox` / `gallery` | Lists filled by callbacks (`getItemCount`, `getItemLabel`, ...). `onAction` takes 3 parameters `(control, id As String, index As Integer)`. |
| `editBox` | A text box. `onChange(control, text As String)`. |
| `box` | A layout helper (`boxStyle="horizontal"` or `"vertical"`). |
| `dynamicMenu` | A menu whose XML is built at run time by `getContent`. Rarely worth the complexity. |

Text attributes: `label` (on the control), `screentip` (bold hover title),
`supertip` (hover body text), `keytip` (Alt shortcut).

## 3. Callback signatures - get these exactly right

Every callback must be a **`Public` Sub in a standard module** (not a form,
class, sheet or ThisWorkbook module). It must exist **once** in the project,
because callbacks resolve by bare name and two Public procedures with the
same name are ambiguous. It must also take the parameters the ribbon passes.

| Attribute | Signature |
|---|---|
| `onLoad` (on `<customUI>`) | `Sub X(ribbon As IRibbonUI)` |
| `onAction` on `button` | `Sub X(control As IRibbonControl)` |
| `onAction` on `toggleButton` / `checkBox` | `Sub X(control As IRibbonControl, pressed As Boolean)` |
| `onAction` on `dropDown` / `gallery` | `Sub X(control As IRibbonControl, id As String, index As Integer)` |
| `getLabel`, `getEnabled`, `getVisible`, `getImage`, `getPressed`, `getScreentip`, `getSupertip`, `getKeytip`, `getItemCount`, `getSelectedItemIndex`, `getText` | `Sub X(control As IRibbonControl, ByRef returnedVal)` |
| `getItemLabel`, `getItemID`, `getItemImage` | `Sub X(control As IRibbonControl, index As Integer, ByRef returnedVal)` |
| `onChange` on `editBox` / `comboBox` | `Sub X(control As IRibbonControl, text As String)` |
| `loadImage` (on `<customUI>`) | `Sub X(imageId As String, ByRef image)` |

A trailing `Optional` parameter is tolerated, because the ribbon passes fewer
arguments and VBA fills in the default. `review_addin.py` accepts that.

**The rule that bites most:** *every `onAction` must exist as a Public Sub
with the right signature, or the button silently does nothing.* Sometimes
Excel shows "Cannot run the macro", but it never says which XML attribute
caused it. `review_addin.py` checks every callback against the code. Run it
after every ribbon edit.

`IRibbonControl` and `IRibbonUI` come from the *Microsoft Office xx.0 Object
Library*, which every Excel VBA project references by default.

### Thin callbacks

Keep each callback to a line or two that hands off to a plain procedure:

```vba
Public Sub Ribbon_InsertNextItem(control As IRibbonControl)
    If Len(InsertNextItem(ActiveCell)) = 0 Then MsgBox "The list is empty."
End Sub

Public Function InsertNextItem(ByVal target As Range) As String   ' testable, no ribbon types
```

There are two reasons. `Application.Run` (and so any test harness) cannot
create an `IRibbonControl`, so logic trapped inside a callback cannot be
tested directly. And a procedure that takes a `Range` works from the Immediate
window, from other macros and from tests alike.

### One shared callback per kind

`getLabel="Ribbon_GetLabel"` on several controls is fine. Branch on
`control.ID` with `Select Case`. Return a visible default in `Case Else`
(for example the control id) so a missing case shows on screen instead of
failing silently.

## 4. Dynamic controls: onLoad and invalidation

Excel calls every `get*` callback **once**, when the control first appears,
and then caches the answer. It asks again only when VBA calls
`Invalidate` (all controls) or `InvalidateControl "id"` (one control) on the
`IRibbonUI` object. The only way to obtain that object is the `onLoad`
callback:

```vba
Private mRibbon As IRibbonUI

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Public Sub RefreshRibbon(Optional ByVal controlId As String = "")
    If mRibbon Is Nothing Then Exit Sub       ' lost after a VBA reset - degrade quietly
    On Error Resume Next
    If Len(controlId) = 0 Then mRibbon.Invalidate Else mRibbon.InvalidateControl controlId
End Sub
```

Points that are easy to miss:
- **No `onLoad` attribute, no refresh, ever.** Labels and enabled states freeze
  at whatever they were when Excel started. Nothing errors.
- **The reference is lost whenever VBA resets**: an unhandled error anywhere in
  the project, or pressing Stop or editing code in the VBA editor. From then
  until Excel restarts, `mRibbon` is `Nothing`. Guard every use with
  `Is Nothing`. Make `get*` callbacks and Application event handlers
  incapable of raising (use `On Error Resume Next` and return a safe default),
  because an error inside one is exactly what resets VBA. Recovery tricks that
  store the object pointer on a sheet exist, but they rely on undocumented
  memory copying, and crashing Excel is worse than a stale label.
- **Nothing refreshes itself on workbook switch.** If a control depends on the
  active workbook (enabled only when a workbook is open, or a list of that
  workbook's names), listen to Application events in `ThisWorkbook`
  (`templates/ThisWorkbook.txt`) and call `RefreshRibbon` from them.
- A `dropDown`'s `onAction` fires only when the selection **changes**.
  Re-picking the selected item runs no code.
- `get*` callbacks run often (`getItemLabel` once per item). Cache the data in
  a module variable and rebuild the cache when you invalidate. Do not re-read
  a sheet on every call.

## 5. Icons

| Way | XML | When |
|---|---|---|
| Built-in | `imageMso="FileSave"` | Thousands available. Find names in the RibbonX Editor's icon gallery or Microsoft's Office icon lists. **Case-sensitive.** An unknown name is not an error: the icon is simply blank. `verify_vba.py` checks every name against the installed Office. |
| Embedded custom | `image="myIcon"` | A PNG stored in the file. In the RibbonX Editor use *Insert Icons*, which writes `customUI/images/myIcon.png` and a relationship with `Id="myIcon"` in `customUI/_rels/customUI14.xml.rels`. `scripts/inject_ribbon.py --image myIcon=my.png` does the same. Use 32x32 for large buttons and 16x16 for normal ones, with transparent backgrounds. |
| Changes with state | `getImage="Ribbon_X"` | The callback may return an **imageMso name as a String** (simplest), or an `IPictureDisp` picture. |
| Loaded from disk | `loadImage="Ribbon_LoadImage"` on `<customUI>`, then `image="name"` | Excel calls `loadImage(imageId, ByRef image)` for each name. Typically `Set image = LoadPicture(ThisWorkbook.Path & "\icons\" & imageId & ".bmp")`. `LoadPicture` cannot read PNG, and the icon files must ship beside the add-in, so embedding is usually the better choice. |

A control takes **one** of `imageMso`, `image` or `getImage`, never two.

## 6. Mutually exclusive attributes

A property and its callback cannot both appear on one control. If they do,
the XML fails schema validation and **the whole tab disappears**:

`label`/`getLabel`, `enabled`/`getEnabled`, `visible`/`getVisible`,
`screentip`/`getScreentip`, `supertip`/`getSupertip`, `keytip`/`getKeytip`,
`size`/`getSize`, `showLabel`/`getShowLabel`, `showImage`/`getShowImage`,
`description`/`getDescription`, and `imageMso`/`image`/`getImage`.

## 7. Editing the ribbon - two routes

**Office RibbonX Editor (default, visible).** Free and open source
(search for "Office RibbonX Editor"; it succeeds the older "Custom UI
Editor").
1. Close the workbook in Excel.
2. *File > Open* the `.xlsm`. If there is no `customUI14.xml` yet, use
   *Insert > Office 2010+ Custom UI Part*.
3. Paste the XML or snippet at the insertion point the snippet's comment
   describes.
4. Click **Validate**. This catches malformed XML, duplicate ids and
   exclusive-attribute clashes. It does **not** check that callbacks exist.
5. *Generate Callbacks* shows correctly-typed empty Subs for every callback,
   which helps you get signatures right. Copy only the ones you do not
   already have.
6. Save, reopen in Excel, and look at the tab.

**Scripted.** `python scripts/inject_ribbon.py MyAddIn.xlsm ribbon.xml --out MyAddIn-new.xlsm`.
This writes the part, the `_rels/.rels` relationship and any icons into a new
file. Useful for repeatable builds and tests.

Either way, finish with `python scripts/review_addin.py`. That is the step
that checks the callbacks.

### Writing a snippet for someone else to paste

Put an XML comment at the top saying exactly where it goes. Name the element
to search for and say "paste immediately above / below / before `</group>`".
Name each callback and the module that holds it. **Never write two hyphens in
a row inside an XML comment**, because that is illegal XML and some editors
cut the comment short or refuse the file. Check it before handing it over:

```
python scripts/review_addin.py MyAddIn.xlsm --xml "ribbon - new group.xml" --bas-dir build --pending
```

A snippet's ids are checked against the ribbon it will be pasted into.

## 8. Debugging a ribbon

1. Tick *File > Options > Advanced > General > Show add-in user interface
   errors* while developing. Excel then reports XML errors at load, with a
   line number, instead of silently dropping the tab.
2. Run `review_addin.py`. It finds most problems without opening Excel.
3. Tab missing: work through the relationship, namespace, well-formedness,
   duplicate ids and exclusive pairs, in that order.
4. Button dead: check the callback name, `Public`, standard module, number of
   parameters, and whether the name is duplicated.
5. Label or enabled state stale: check `onLoad` is present, `mRibbon` is not
   `Nothing`, and you called `InvalidateControl` with the right id.
6. If the cause is still unclear, write a throwaway test: a one-button XML
   and a one-line callback, then add things back until it breaks. Experiment
   beats inference here.
