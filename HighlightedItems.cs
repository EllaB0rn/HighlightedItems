using System.Windows.Forms;
using HighlightedItems.Utils;
using ExileCore;
using ExileCore.PoEMemory.Elements;
using ExileCore.PoEMemory.Elements.InventoryElements;
using ExileCore.PoEMemory.MemoryObjects;
using ExileCore.PoEMemory;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExileCore.Shared.Enums;
using ImGuiNET;
using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using ExileCore.PoEMemory.Components;
using ExileCore.Shared;
using ExileCore.Shared.Helpers;
using ItemFilterLibrary;
using SharpDX;
using Vector2 = System.Numerics.Vector2;

namespace HighlightedItems;

public class HighlightedItems : BaseSettingsPlugin<Settings>
{
    private SyncTask<bool> _currentOperation;
    private string _customStashFilter = "";
    private string _customInventoryFilter = "";
    private BestiaryPanel _cachedBestiaryPanel;
    private Element _cachedBestiaryElement;
    private long _lastBestiaryPanelCheck;
    private Dictionary<long, bool> _cachedInventoryItemsSaturation;
    private long _lastInventorySaturationCheck;
    private Task _saturationUpdateTask;

    // Constants
    private const int MAX_INVENTORY_ITEMS = 60;
    private const int BESTIARY_CATEGORIES = 12;
    private const int BEASTS_PER_BATCH = 9;
    private const int WORK_PAUSE_INTERVAL_MS = 10000;
    private const int WORK_PAUSE_DURATION_MS = 1000;

    private record QueryOrException(ItemQuery Query, Exception Exception);

    private readonly ConditionalWeakTable<string, QueryOrException> _queries = [];

    private bool MoveCancellationRequested => Settings.CancelWithRightMouseButton && (Control.MouseButtons & MouseButtons.Right) != 0;
    private IngameState InGameState => GameController.IngameState;
    private SharpDX.Vector2 WindowOffset => GameController.Window.GetWindowRectangleTimeCache.TopLeft;

    public override bool Initialise()
    {
        Graphics.InitImage(Path.Combine(DirectoryFullName, "images\\pick.png").Replace('\\', '/'), false);
        Graphics.InitImage(Path.Combine(DirectoryFullName, "images\\pickL.png").Replace('\\', '/'), false);

        return true;
    }

    public override void AreaChange(AreaInstance area)
    {
        _mouseStateForRect.Clear();
    }

    public override void DrawSettings()
    {
        base.DrawSettings();
        DrawIgnoredCellsSettings();
    }

    private Predicate<Entity> GetPredicate(string windowTitle, ref string filterText, Vector2 defaultPosition)
    {
        if (!Settings.ShowCustomFilterWindow) return null;
        Settings.SavedFilters ??= [];
        ImGui.SetNextWindowPos(defaultPosition, ImGuiCond.FirstUseEver);
        if (ImGui.Begin(windowTitle, ImGuiWindowFlags.NoDecoration | ImGuiWindowFlags.AlwaysAutoResize))
        {
            ImGui.InputTextWithHint("##input", "Filter using IFL syntax", ref filterText, 2000);
            Predicate<Entity> returnValue = null;
            if (!string.IsNullOrWhiteSpace(filterText))
            {
                ImGui.SameLine();
                if (ImGui.Button("Clear"))
                {
                    filterText = "";
                    return null;
                }

                if (!Settings.SavedFilters.Contains(filterText))
                {
                    ImGui.SameLine();
                    if (ImGui.Button("Save"))
                    {
                        Settings.SavedFilters.Add(filterText);
                    }
                }

                var (query, exception) = _queries.GetValue(filterText, s =>
                {
                    try
                    {
                        var itemQuery = ItemQuery.Load(s);
                        if (itemQuery.FailedToCompile)
                        {
                            return new QueryOrException(null, new Exception(itemQuery.Error));
                        }

                        return new QueryOrException(itemQuery, null);
                    }
                    catch (Exception ex)
                    {
                        return new QueryOrException(null, ex);
                    }
                })!;

                if (exception != null)
                {
                    ImGui.TextUnformatted($"{exception.Message}");
                }
                else
                {
                    returnValue = s =>
                    {
                        try
                        {
                            return query.CompiledQuery(new ItemData(s, GameController));
                        }
                        catch (Exception ex)
                        {
                            DebugWindow.LogError($"Failed to match item: {ex}");
                            return false;
                        }
                    };
                }
            }

            // ReSharper disable once AssignmentInConditionalExpression
            if (Settings.SavedFilters.Any() && Settings.UsePopupForFilterSelector
                    ? Settings.OpenSavedFilterList = ImGui.BeginPopupContextItem("saved_filter_popup")
                    : Settings.OpenSavedFilterList = ImGui.TreeNodeEx("Saved filters",
                        Settings.OpenSavedFilterList
                            ? ImGuiTreeNodeFlags.DefaultOpen | ImGuiTreeNodeFlags.NoTreePushOnOpen
                            : ImGuiTreeNodeFlags.NoTreePushOnOpen))
            {
                foreach (var (savedFilter, index) in Settings.SavedFilters.Select((x, i) => (x, i)).ToList())
                {
                    ImGui.PushID($"saved{index}");
                    if (ImGui.Button("Load"))
                    {
                        filterText = savedFilter;
                    }

                    ImGui.SameLine();
                    if (ImGui.Button("Delete"))
                    {
                        if (ImGui.IsKeyDown(ImGuiKey.ModShift))
                        {
                            Settings.SavedFilters.Remove(savedFilter);
                        }
                    }
                    else if (ImGui.IsItemHovered())
                    {
                        ImGui.SetTooltip("Hold Shift");
                    }

                    ImGui.SameLine();
                    ImGui.TextUnformatted(savedFilter);

                    ImGui.PopID();
                }

                if (Settings.UsePopupForFilterSelector)
                {
                    ImGui.EndPopup();
                }
                else
                {
                    ImGui.TreePop();
                }
            }

            if (Settings.UsePopupForFilterSelector)
            {
                if (ImGui.Button("Open Saved Filters"))
                {
                    ImGui.OpenPopup("saved_filter_popup");
                }
            }

            ImGui.End();
            return returnValue;
        }

        return null;
    }

    public override void Render()
    {
        if (_currentOperation != null)
        {
            DebugWindow.LogMsg("Running the inventory dump procedure...");
            TaskUtils.RunOrRestart(ref _currentOperation, () => null);
            if (_itemsToMove is { Count: > 0 } itemsToMove)
            {
                foreach (var (rect, color) in itemsToMove.Skip(1).Select(x => (x, Settings.CustomFilterFrameColor)).Prepend((itemsToMove[0], Color.Green)))
                {
                    Graphics.DrawFrame(rect.TopLeft.ToVector2Num(), rect.BottomRight.ToVector2Num(), color, Settings.CustomFilterFrameThickness);
                }
            }
            return;
        }

        if (!Settings.Enable)
            return;

        var (inventory, rectElement) = (InGameState.IngameUi.StashElement, InGameState.IngameUi.GuildStashElement) switch
        {
            ({ IsVisible: true, VisibleStash: { InventoryUIElement: { } invRect } visibleStash }, _) => (visibleStash, invRect),
            (_, { IsVisible: true, VisibleStash: { InventoryUIElement: { } invRect } visibleStash }) => (visibleStash, invRect),
            _ => (null, null)
        };

        const float buttonSize = 37;
        var highlightedItemsFound = false;
        if (inventory != null)
        {
            var stashRect = rectElement.GetClientRectCache;
            var (itemFilter, isCustomFilter) = GetPredicate("Custom stash filter", ref _customStashFilter, stashRect.BottomLeft.ToVector2Num()) is { } customPredicate
                ? ((Predicate<NormalInventoryItem>)(s => customPredicate(s.Item)), true)
                : (s => s.isHighlighted != Settings.InvertSelection.Value, false);

            //Determine Stash Pickup Button position and draw
            var buttonPos = Settings.UseCustomMoveToInventoryButtonPosition
                ? Settings.CustomMoveToInventoryButtonPosition
                : stashRect.BottomRight.ToVector2Num() + new Vector2(-43, 10);
            var buttonRect = new SharpDX.RectangleF(buttonPos.X, buttonPos.Y, buttonSize, buttonSize);

            Graphics.DrawImage("pick.png", buttonRect);

            var highlightedItems = GetHighlightedItems(inventory, itemFilter);
            highlightedItemsFound = highlightedItems.Any();
            int? stackSizes = 0;
            foreach (var item in highlightedItems)
            {
                stackSizes += item.Item?.GetComponent<Stack>()?.Size;
                if (isCustomFilter)
                {
                    var rect = item.GetClientRectCache;
                    var deflateFactor = Settings.CustomFilterBorderDeflation / 200.0;
                    var deflateWidth = (int)(rect.Width * deflateFactor + Settings.CustomFilterFrameThickness / 2);
                    var deflateHeight = (int)(rect.Height * deflateFactor + Settings.CustomFilterFrameThickness / 2);
                    rect.Inflate(-deflateWidth, -deflateHeight);

                    var topLeft = rect.TopLeft.ToVector2Num();
                    var bottomRight = rect.BottomRight.ToVector2Num();
                    Graphics.DrawFrame(topLeft, bottomRight, Settings.CustomFilterFrameColor, Settings.CustomFilterBorderRounding, Settings.CustomFilterFrameThickness, 0);
                }
            }

            var countText = Settings.ShowStackSizes && highlightedItems.Count != stackSizes && stackSizes != null
                ? Settings.ShowStackCountWithSize
                    ? $"{stackSizes} / {highlightedItems.Count}"
                    : $"{stackSizes}"
                : $"{highlightedItems.Count}";

            var countPos = new Vector2(buttonRect.Left - 2, buttonRect.Center.Y - 11);
            Graphics.DrawText($"{countText}", countPos with { Y = countPos.Y + 2 }, SharpDX.Color.Black, FontAlign.Right);
            Graphics.DrawText($"{countText}", countPos with { X = countPos.X - 2 }, SharpDX.Color.White, FontAlign.Right);

            if (IsButtonPressed(buttonRect) ||
                Input.IsKeyDown(Settings.MoveToInventoryHotkey.Value))
            {
                _currentOperation = MoveItemsToInventoryWithRetry(inventory, itemFilter);
            }
        }
        else
        {
            if (Settings.ResetCustomFilterOnPanelClose)
            {
                _customStashFilter = "";
            }
        }

        var inventoryPanel = InGameState.IngameUi.InventoryPanel;
        if (inventoryPanel.IsVisible)
        {
            var inventoryRect = inventoryPanel[2].GetClientRectCache;

            var (itemFilter, isCustomFilter) = GetPredicate("Custom inventory filter", ref _customInventoryFilter, inventoryRect.BottomLeft.ToVector2Num()) is { } customPredicate
                ? (customPredicate, true)
                : (_ => true, false);

            if (Settings.DumpButtonEnable && IsStashTargetOpened)
            {
                //Determine Inventory Pickup Button position and draw
                var buttonPos = Settings.UseCustomMoveToStashButtonPosition
                    ? Settings.CustomMoveToStashButtonPosition
                    : inventoryRect.TopLeft.ToVector2Num() + new Vector2(buttonSize / 2, -buttonSize);
                var buttonRect = new SharpDX.RectangleF(buttonPos.X, buttonPos.Y, buttonSize, buttonSize);

                if (isCustomFilter)
                {
                    foreach (var item in GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems
                        .Where(x => itemFilter(x.Item)))
                    {
                        var rect = item.GetClientRect();
                        var deflateFactor = Settings.CustomFilterBorderDeflation / 200.0;
                        var deflateWidth = (int)(rect.Width * deflateFactor + Settings.CustomFilterFrameThickness / 2);
                        var deflateHeight = (int)(rect.Height * deflateFactor + Settings.CustomFilterFrameThickness / 2);
                        rect.Inflate(-deflateWidth, -deflateHeight);

                        var topLeft = rect.TopLeft.ToVector2Num();
                        var bottomRight = rect.BottomRight.ToVector2Num();
                        Graphics.DrawFrame(topLeft, bottomRight, Settings.CustomFilterFrameColor, Settings.CustomFilterBorderRounding, Settings.CustomFilterFrameThickness, 0);
                    }
                }

                Graphics.DrawImage("pickL.png", buttonRect);
                if (IsButtonPressed(buttonRect) ||
                    Input.IsKeyDown(Settings.MoveToStashHotkey.Value) ||
                    Settings.UseMoveToInventoryAsMoveToStashWhenNoHighlights &&
                    !highlightedItemsFound &&
                    Input.IsKeyDown(Settings.MoveToInventoryHotkey.Value))
                {
                    _currentOperation = MoveItemsToStashWithRetry(itemFilter);
                }
            }
        }
        else
        {
            if (Settings.ResetCustomFilterOnPanelClose)
            {
                _customInventoryFilter = "";
            }
        }

        // Bestiary panel handling - check hotkeys FIRST, then search for panel only if needed
        // This prevents expensive panel search every frame
        bool bestiaryHotkeyPressed = false;

        if (Settings.BestiaryToInventoryWithCheckHotkey.Value != Keys.None &&
            Input.IsKeyDown(Settings.BestiaryToInventoryWithCheckHotkey.Value))
        {
            bestiaryHotkeyPressed = true;
            DebugWindow.LogMsg("HighlightedItems: Moving beasts to inventory (with check - hotkey)");
            _currentOperation = MoveBeastsToInventoryWithCheck();
        }
        else if (Settings.BestiaryToInventoryForceHotkey.Value != Keys.None &&
                 Input.IsKeyDown(Settings.BestiaryToInventoryForceHotkey.Value))
        {
            bestiaryHotkeyPressed = true;
            DebugWindow.LogMsg("HighlightedItems: Moving beasts to inventory (force mode)");
            _currentOperation = MoveBeastsToInventoryForce();
        }
        else if (Settings.BestiaryDebugHotkey.Value != Keys.None &&
                 Input.IsKeyDown(Settings.BestiaryDebugHotkey.Value))
        {
            bestiaryHotkeyPressed = true;
            DebugWindow.LogMsg("HighlightedItems: Bestiary debug");
            _currentOperation = DebugBestiaryPanel();
        }
    }

    // Bestiary panel classes
    private class BestiaryPanel : Element
    {
        // BestiaryPanel directly contains 12 category children
        public List<CapturedBeast> GetAllBeasts()
        {
            var allBeasts = new List<CapturedBeast>();

            // Iterate through all 12 beast type categories
            for (int i = 0; i < 12; i++)
            {
                var categoryElement = GetChildAtIndex(i);
                if (categoryElement == null || !categoryElement.IsVisible) continue;

                // Each category has 2 children, Child[1] contains the beasts
                var beastsContainer = categoryElement.GetChildAtIndex(1);
                if (beastsContainer == null) continue;

                // Get all beast children
                foreach (var beastChild in beastsContainer.Children)
                {
                    if (beastChild != null && beastChild.IsVisible)
                    {
                        allBeasts.Add(GetObject<CapturedBeast>(beastChild.Address));
                    }
                }
            }

            return allBeasts;
        }

        // Get beasts from a specific category (0-11)
        // maxBeasts: limit how many beasts to retrieve (0 = all, default = 9 for performance)
        public List<CapturedBeast> GetBeastsFromCategory(int categoryIndex, int maxBeasts = 9)
        {
            var beasts = new List<CapturedBeast>();

            if (categoryIndex < 0 || categoryIndex >= 12) return beasts;

            var categoryElement = GetChildAtIndex(categoryIndex);
            if (categoryElement == null || !categoryElement.IsVisible) return beasts;

            // Each category has 2 children, Child[1] contains the beasts
            var beastsContainer = categoryElement.GetChildAtIndex(1);
            if (beastsContainer == null) return beasts;

            // Get beast children from this category (limited to maxBeasts for performance)
            int count = 0;
            foreach (var beastChild in beastsContainer.Children)
            {
                if (beastChild != null && beastChild.IsVisible)
                {
                    beasts.Add(GetObject<CapturedBeast>(beastChild.Address));
                    count++;

                    // Stop early if we have enough beasts (optimization for large collections)
                    if (maxBeasts > 0 && count >= maxBeasts)
                    {
                        break;
                    }
                }
            }

            return beasts;
        }
    }

    private class CapturedBeast : Element
    {
        public string DisplayName => Tooltip?.GetChildAtIndex(1)?.GetChildAtIndex(0)?.Text?.Replace("-", "").Trim();
    }

    private async SyncTask<bool> DebugBestiaryPanel()
    {
        DebugWindow.LogMsg("=== BESTIARY DEBUG START ===");

        try
        {
            var bestiaryPanel = GetBestiaryPanel();
            if (bestiaryPanel == null)
            {
                DebugWindow.LogMsg("BestiaryPanel is NULL");
                return false;
            }

            DebugWindow.LogMsg($"BestiaryPanel found, Address: {bestiaryPanel.Address:X}");
            DebugWindow.LogMsg($"BestiaryPanel.ChildCount: {bestiaryPanel.ChildCount}");

            var beasts = bestiaryPanel.GetAllBeasts();
            DebugWindow.LogMsg($"Total beasts found: {beasts.Count}");

            int visibleBeasts = 0;
            foreach (var beast in beasts)
            {
                if (beast != null && beast.IsVisible)
                {
                    visibleBeasts++;
                    var rect = beast.GetClientRect();
                    DebugWindow.LogMsg($"Beast #{visibleBeasts}: {beast.DisplayName ?? "Unknown"}, Pos: ({rect.X}, {rect.Y})");

                    if (visibleBeasts >= 5) // Show only first 5 for brevity
                    {
                        DebugWindow.LogMsg($"... and {beasts.Count - visibleBeasts} more beasts");
                        break;
                    }
                }
            }

            DebugWindow.LogMsg($"Total visible beasts: {visibleBeasts}");

            // Check inventory space
            var inventory = GameController.IngameState.ServerData.PlayerInventories[0].Inventory;
            int currentCount = (int)inventory.CountItems;
            int maxItems = 60;
            int freeSlots = maxItems - currentCount;

            DebugWindow.LogMsg($"Inventory: {currentCount}/{maxItems}, Free slots: {freeSlots}");
            DebugWindow.LogMsg($"Would move: {Math.Min(visibleBeasts, freeSlots)} beasts");
        }
        catch (Exception ex)
        {
            DebugWindow.LogMsg($"Error in DebugBestiaryPanel: {ex.Message}");
            DebugWindow.LogMsg($"Stack trace: {ex.StackTrace}");
        }

        DebugWindow.LogMsg("=== BESTIARY DEBUG END ===");
        return true;
    }

    private BestiaryPanel GetBestiaryPanelCached()
    {
        var now = Environment.TickCount64;

        // If we have a cached panel, check if it's still valid (cache for 1 second)
        if (_cachedBestiaryPanel != null && _cachedBestiaryElement != null && (now - _lastBestiaryPanelCheck) < 1000)
        {
            try
            {
                // IMPORTANT: Check visibility - if the element is not visible, the bestiary is closed
                if (!_cachedBestiaryElement.IsVisible)
                {
                    _cachedBestiaryPanel = null;
                    _cachedBestiaryElement = null;
                    return null;
                }

                // Check if panel still has valid structure
                if (_cachedBestiaryElement.Address != 0 && _cachedBestiaryElement.ChildCount == 12)
                {
                    return _cachedBestiaryPanel;
                }
            }
            catch
            {
                // Panel became invalid, clear cache
                _cachedBestiaryPanel = null;
                _cachedBestiaryElement = null;
            }
        }

        // If we cached null (panel not found), wait 100ms before searching again
        // This provides quick detection while preventing excessive searches
        if (_cachedBestiaryPanel == null && (now - _lastBestiaryPanelCheck) < 100)
        {
            return null;
        }

        // Cache expired or invalid, do a new search
        _lastBestiaryPanelCheck = now;
        _cachedBestiaryPanel = GetBestiaryPanel();
        return _cachedBestiaryPanel;
    }

    private BestiaryPanel GetBestiaryPanel()
    {
        try
        {
            var ui = InGameState.IngameUi;
            if (ui == null) return null;

            // Search for element with 12 children starting from child50
            var child50 = ui.GetChildAtIndex(50);
            if (child50 == null) return null;

            var panelWith12Children = FindElementWith12Children(child50, 0, 12);
            if (panelWith12Children != null)
            {
                // IMPORTANT: Verify this is actually the bestiary panel by checking visibility
                // The bestiary panel should only be visible when the bestiary menu is open
                if (!panelWith12Children.IsVisible)
                {
                    _cachedBestiaryElement = null;
                    return null;
                }

                // Check that the element is in the LEFT side of the screen (bestiary panel position)
                // The bestiary panel should be on the left side, not in the top-left corner
                var rect = panelWith12Children.GetClientRectCache;

                // Bestiary panel should be:
                // - At least 400 pixels wide and 500 pixels tall (it's a large panel)
                // - Positioned on the left side of the screen (X < 600)
                // - Not in the very top corner (Y > 50)
                if (rect.Width < 400 || rect.Height < 500 || rect.Left > 600 || rect.Top < 50)
                {
                    _cachedBestiaryElement = null;
                    return null;
                }

                // Additional check: verify the panel has the expected structure
                // Each of the 12 children should have 2 children (header + beasts container)
                bool hasValidStructure = true;
                for (int i = 0; i < Math.Min(12, panelWith12Children.ChildCount); i++)
                {
                    var category = panelWith12Children.GetChildAtIndex(i);
                    if (category == null || category.ChildCount != 2)
                    {
                        hasValidStructure = false;
                        break;
                    }
                }

                if (!hasValidStructure)
                {
                    _cachedBestiaryElement = null;
                    return null;
                }

                _cachedBestiaryElement = panelWith12Children;
                return ui.GetObject<BestiaryPanel>(panelWith12Children.Address);
            }

            _cachedBestiaryElement = null;
            return null;
        }
        catch
        {
            _cachedBestiaryElement = null;
            return null;
        }
    }

    private dynamic FindElementWith12Children(dynamic element, int currentDepth, int maxDepth)
    {
        if (element == null || currentDepth > maxDepth) return null;

        // Check if this element has exactly 12 children
        if (element.ChildCount == 12)
        {
            // Verify it's the bestiary panel by checking if children have the expected structure
            var firstChild = element.GetChildAtIndex(0);
            if (firstChild != null && firstChild.ChildCount == 2)
            {
                return element;
            }
        }

        // Recursively search children
        for (int i = 0; i < element.ChildCount && i < 20; i++)
        {
            var child = element.GetChildAtIndex(i);
            var result = FindElementWith12Children(child, currentDepth + 1, maxDepth);
            if (result != null) return result;
        }

        return null;
    }

    // Helper method to get current inventory count
    private int GetInventoryCount() => (int)GameController.IngameState.ServerData.PlayerInventories[0].Inventory.CountItems;

    // Helper method to check if inventory is full
    private bool IsInventoryFull() => GetInventoryCount() >= MAX_INVENTORY_ITEMS;

    // Helper method to get remaining inventory slots
    private int GetRemainingInventorySlots() => MAX_INVENTORY_ITEMS - GetInventoryCount();

    private async SyncTask<bool> MoveBeastsToInventoryWithCheck()
    {
        DebugWindow.LogMsg("HighlightedItems: MoveBeastsToInventoryWithCheck started");

        // Wait for mouse to be released
        while (Control.MouseButtons == MouseButtons.Left || MoveCancellationRequested)
        {
            if (MoveCancellationRequested) return false;
            await TaskUtils.NextFrame();
        }

        // Check if Bestiary panel is open
        var bestiaryPanel = GetBestiaryPanelCached();
        if (bestiaryPanel == null)
        {
            DebugWindow.LogMsg("HighlightedItems: Bestiary panel is not open");
            return false;
        }

        // Open inventory if needed
        if (!InGameState.IngameUi.InventoryPanel.IsVisible)
        {
            DebugWindow.LogMsg("HighlightedItems: Opening inventory...");
            Keyboard.KeyDown(Keys.I);
            Keyboard.KeyUp(Keys.I);
            await Wait(TimeSpan.FromMilliseconds(150), false);

            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                DebugWindow.LogMsg("HighlightedItems: Failed to open inventory");
                return false;
            }
        }

        // Check inventory space
        int totalFreeSlots = GetRemainingInventorySlots();
        if (totalFreeSlots <= 0)
        {
            DebugWindow.LogMsg("HighlightedItems: Inventory full, cannot move beasts");
            return false;
        }

        DebugWindow.LogMsg($"HighlightedItems: Starting beast transfer. Free slots: {totalFreeSlots}");

        // Save initial mouse position and press Ctrl
        _prevMousePos = Mouse.GetCursorPosition();
        Keyboard.KeyDown(Keys.LControlKey);
        await Wait(KeyDelay, true);

        int totalMoved = 0;
        int waitTimeout = Settings.FastMode.Value ? 300 : 600;
        int maxRetries = Settings.RetryMissedItems.Value ? Settings.MaxRetryAttempts.Value : 0;
        int currentCategory = 0;
        bool isFirstBatch = true;
        var workTimer = Stopwatch.StartNew();

        while (totalMoved < totalFreeSlots && currentCategory < BESTIARY_CATEGORIES)
        {
            if (MoveCancellationRequested)
            {
                await StopMovingItems();
                return totalMoved > 0;
            }

            // Pause after continuous work
            if (workTimer.ElapsedMilliseconds >= WORK_PAUSE_INTERVAL_MS)
            {
                await Wait(TimeSpan.FromMilliseconds(WORK_PAUSE_DURATION_MS), false);
                workTimer.Restart();
            }

            // Check if inventory is full
            if (IsInventoryFull())
            {
                DebugWindow.LogMsg("HighlightedItems: Inventory full, stopping");
                break;
            }

            bestiaryPanel = GetBestiaryPanelCached();
            if (bestiaryPanel == null)
            {
                await StopMovingItems();
                return totalMoved > 0;
            }

            // Get beasts from current category
            var beasts = bestiaryPanel.GetBeastsFromCategory(currentCategory, BEASTS_PER_BATCH);
            if (beasts == null || beasts.Count == 0)
            {
                currentCategory++;
                continue;
            }

            beasts.Reverse();

            // Calculate how many beasts to move
            int beastsToTake = Math.Min(BEASTS_PER_BATCH, Math.Min(beasts.Count, GetRemainingInventorySlots()));
            if (beastsToTake <= 0) break;

            var beastsToMove = beasts.Take(beastsToTake).ToList();
            int inventoryCountBefore = GetInventoryCount();
            int retryAttempt = 0;
            int actuallyMoved = 0;

            while (retryAttempt <= maxRetries)
            {
                if (isFirstBatch)
                {
                    await Wait(TimeSpan.FromMilliseconds(30), true);
                    isFirstBatch = false;
                }

                if (!await MoveBeastsInternal(beastsToMove, bestiaryPanel))
                {
                    await StopMovingItems();
                    return totalMoved > 0;
                }

                await WaitForInventoryChange(inventoryCountBefore, waitTimeout);

                int inventoryCountAfter = GetInventoryCount();
                actuallyMoved = inventoryCountAfter - inventoryCountBefore;

                if (IsInventoryFull())
                {
                    totalMoved += actuallyMoved;
                    await StopMovingItems();
                    return true;
                }

                if (actuallyMoved >= beastsToMove.Count) break;

                var beastsAfter = bestiaryPanel.GetBeastsFromCategory(currentCategory, 1);
                if (beastsAfter == null || beastsAfter.Count == 0) break;

                if (retryAttempt < maxRetries && actuallyMoved < beastsToMove.Count)
                {
                    int missedBeasts = beastsToMove.Count - actuallyMoved;
                    var beastsForRetry = bestiaryPanel.GetBeastsFromCategory(currentCategory, Math.Min(BEASTS_PER_BATCH, missedBeasts));
                    beastsForRetry.Reverse();
                    beastsToMove = beastsForRetry.Take(missedBeasts).ToList();
                    inventoryCountBefore = inventoryCountAfter;
                    retryAttempt++;
                }
                else break;
            }

            totalMoved += actuallyMoved;
            if (IsInventoryFull() || totalMoved >= totalFreeSlots) break;
            if (actuallyMoved == 0)
            {
                currentCategory++;
                continue;
            }

            var remainingInCategory = bestiaryPanel.GetBeastsFromCategory(currentCategory, 1);
            if (remainingInCategory == null || remainingInCategory.Count == 0) currentCategory++;

            await Wait(TimeSpan.FromMilliseconds(Settings.FastMode.Value ? 20 : 100), false);
        }

        DebugWindow.LogMsg($"HighlightedItems: Finished. Total moved: {totalMoved}");
        await StopMovingItems();
        return true;
    }

    private async SyncTask<bool> MoveBeastsToInventoryForce()
    {
        DebugWindow.LogMsg("HighlightedItems: MoveBeastsToInventoryForce started");

        // Wait for mouse to be released
        while (Control.MouseButtons == MouseButtons.Left || MoveCancellationRequested)
        {
            if (MoveCancellationRequested) return false;
            await TaskUtils.NextFrame();
        }

        // Check if Bestiary panel is open
        var bestiaryPanel = GetBestiaryPanelCached();
        if (bestiaryPanel == null)
        {
            DebugWindow.LogMsg("HighlightedItems: Bestiary panel is not open");
            return false;
        }

        // Open inventory if needed
        if (!InGameState.IngameUi.InventoryPanel.IsVisible)
        {
            DebugWindow.LogMsg("HighlightedItems: Opening inventory...");
            Keyboard.KeyDown(Keys.I);
            Keyboard.KeyUp(Keys.I);
            await Wait(TimeSpan.FromMilliseconds(150), false);

            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                DebugWindow.LogMsg("HighlightedItems: Failed to open inventory");
                return false;
            }
        }

        // Save initial mouse position and press Ctrl
        _prevMousePos = Mouse.GetCursorPosition();
        Keyboard.KeyDown(Keys.LControlKey);
        await Wait(KeyDelay, true);

        int totalMoved = 0;
        int currentCategory = 0;
        bool isFirstBatch = true;
        var workTimer = Stopwatch.StartNew();

        while (currentCategory < BESTIARY_CATEGORIES)
        {
            if (MoveCancellationRequested)
            {
                await StopMovingItems();
                return totalMoved > 0;
            }

            // Pause after continuous work
            if (workTimer.ElapsedMilliseconds >= WORK_PAUSE_INTERVAL_MS)
            {
                await Wait(TimeSpan.FromMilliseconds(WORK_PAUSE_DURATION_MS), false);
                workTimer.Restart();
            }

            bestiaryPanel = GetBestiaryPanelCached();
            if (bestiaryPanel == null)
            {
                await StopMovingItems();
                return totalMoved > 0;
            }

            // Get beasts from current category
            var beasts = bestiaryPanel.GetBeastsFromCategory(currentCategory, BEASTS_PER_BATCH);
            if (beasts == null || beasts.Count == 0)
            {
                currentCategory++;
                continue;
            }

            int beastsBeforeBatch = beasts.Count;
            beasts.Reverse();

            int beastsToTake = Math.Min(BEASTS_PER_BATCH, beasts.Count);
            var beastsToMove = beasts.Take(beastsToTake).ToList();

            if (isFirstBatch)
            {
                await Wait(TimeSpan.FromMilliseconds(30), true);
                isFirstBatch = false;
            }

            if (!await MoveBeastsInternal(beastsToMove, bestiaryPanel))
            {
                await StopMovingItems();
                return totalMoved > 0;
            }

            await Wait(TimeSpan.FromMilliseconds(Settings.FastMode.Value ? 30 : 100), false);

            // Check how many beasts are left
            var beastsAfter = bestiaryPanel.GetBeastsFromCategory(currentCategory, BEASTS_PER_BATCH);
            int beastsAfterBatch = beastsAfter?.Count ?? 0;
            int actuallyMoved = beastsBeforeBatch - beastsAfterBatch;

            totalMoved += actuallyMoved;

            if (beastsAfterBatch == 0) currentCategory++;
        }

        DebugWindow.LogMsg($"HighlightedItems: Finished. Total moved: {totalMoved}");
        await StopMovingItems();
        return true;
    }

    private async SyncTask<bool> MoveBeastsInternal(List<CapturedBeast> beasts, BestiaryPanel panel)
    {
        DebugWindow.LogMsg($"HighlightedItems: MoveBeastsInternal started with {beasts.Count} beasts");

        // Note: Ctrl is already pressed by the parent function, mouse position already saved
        int clickedCount = 0;

        for (var i = 0; i < beasts.Count; i++)
        {
            var beast = beasts[i];

            if (MoveCancellationRequested)
            {
                DebugWindow.LogMsg("HighlightedItems: Cancellation requested, stopping");
                return false;
            }

            var bestiaryPanel = GetBestiaryPanelCached();
            if (bestiaryPanel == null)
            {
                DebugWindow.LogMsg("HighlightedItems: Bestiary Panel closed, aborting loop");
                break;
            }

            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                DebugWindow.LogMsg("HighlightedItems: Inventory Panel closed, aborting loop");
                break;
            }

            // Get beast rectangle
            var beastRect = beast.GetClientRect();

            // Check if beast is within reasonable screen bounds
            // Beasts that are scrolled out of view will have Y coordinates beyond visible panel area
            // Allow beasts up to Y<600 (Y=540 and below are visible)
            if (beastRect.X < 50 || beastRect.Y < 50 ||
                beastRect.Width <= 0 || beastRect.Height <= 0 ||
                beastRect.X > 2500 || beastRect.Y >= 600)
            {
                DebugWindow.LogMsg($"HighlightedItems: Skipping beast #{i+1} - outside visible area (pos: {beastRect.X},{beastRect.Y}, size: {beastRect.Width}x{beastRect.Height})");
                continue;
            }

            DebugWindow.LogMsg($"HighlightedItems: Clicking beast #{i+1} at ({beastRect.Center.X}, {beastRect.Center.Y})");
            await MoveItem(beastRect.Center);
            clickedCount++;
        }

        DebugWindow.LogMsg($"HighlightedItems: Clicked {clickedCount} beasts");

        // Note: Don't release Ctrl or restore mouse here - parent function handles that
        return true;
    }

    private async SyncTask<bool> MoveItemsCommonPreamble()
    {
        DebugWindow.LogMsg("HighlightedItems: MoveItemsCommonPreamble - checking mouse state");

        while (Control.MouseButtons == MouseButtons.Left || MoveCancellationRequested)
        {
            if (MoveCancellationRequested)
            {
                DebugWindow.LogMsg("HighlightedItems: Cancellation requested in preamble");
                return false;
            }

            await TaskUtils.NextFrame();
        }

        if (EffectiveIdleMouseDelay == 0)
        {
            DebugWindow.LogMsg("HighlightedItems: No idle mouse delay configured, proceeding");
            return true;
        }

        DebugWindow.LogMsg($"HighlightedItems: Waiting for mouse to be idle ({EffectiveIdleMouseDelay}ms)");
        var mousePos = Mouse.GetCursorPosition();
        var sw = Stopwatch.StartNew();
        await TaskUtils.NextFrame();
        while (true)
        {
            if (MoveCancellationRequested)
            {
                DebugWindow.LogMsg("HighlightedItems: Cancellation requested while waiting for idle mouse");
                return false;
            }

            var newPos = Mouse.GetCursorPosition();
            if (mousePos != newPos)
            {
                mousePos = newPos;
                sw.Restart();
            }
            else if (sw.ElapsedMilliseconds >= EffectiveIdleMouseDelay)
            {
                DebugWindow.LogMsg("HighlightedItems: Mouse idle, proceeding");
                return true;
            }
            else
            {
                await TaskUtils.NextFrame();
            }
        }
    }

    private async SyncTask<bool> WaitForInventoryChange(int initialCount, int maxWaitMs = 2000)
    {
        var sw = Stopwatch.StartNew();
        DebugWindow.LogMsg($"HighlightedItems: WaitForInventoryChange started (initial: {initialCount}, timeout: {maxWaitMs}ms)");

        while (sw.ElapsedMilliseconds < maxWaitMs)
        {
            var currentCount = (int)GameController.IngameState.ServerData.PlayerInventories[0].Inventory.CountItems;
            if (currentCount != initialCount)
            {
                DebugWindow.LogMsg($"HighlightedItems: Inventory changed from {initialCount} to {currentCount} items after {sw.ElapsedMilliseconds}ms");
                return true;
            }
            await TaskUtils.NextFrame();
        }

        DebugWindow.LogMsg($"HighlightedItems: Timeout waiting for inventory change after {sw.ElapsedMilliseconds}ms (still {initialCount} items)");
        return false;
    }

    private async SyncTask<bool> MoveItemsToStashWithRetry(Predicate<Entity> itemFilter)
    {
        int maxRetries = Settings.RetryMissedItems ? Settings.MaxRetryAttempts.Value : 0;

        // Get current inventory items that match the filter
        var allInventoryItems = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems
            .Where(x => !IsInIgnoreCell(x))
            .Where(x => itemFilter(x.Item))
            .OrderBy(x => x.PosX)
            .ThenBy(x => x.PosY)
            .ToList();

        if (allInventoryItems.Count == 0)
        {
            DebugWindow.LogMsg("HighlightedItems: No items to move");
            return true;
        }

        DebugWindow.LogMsg($"HighlightedItems: Clicking all {allInventoryItems.Count} items immediately");

        // Record initial count before moving
        int initialCount = (int)GameController.IngameState.ServerData.PlayerInventories[0].Inventory.CountItems;

        // Move items immediately without saturation check
        bool success = await MoveItemsToStash(allInventoryItems);

        if (!success)
        {
            return false;
        }

        // Start background task to check and retry missed items
        if (maxRetries > 0)
        {
            StartBackgroundRetryTask(allInventoryItems, maxRetries);
        }

        // Wait for inventory to change (but don't block on retries)
        await WaitForInventoryChange(initialCount, 2000);

        return true;
    }

    private Dictionary<long, bool> GetInventoryItemsSaturationCached()
    {
        var now = Environment.TickCount64;

        // Cache for 200ms - fast enough to catch changes, slow enough to avoid overhead
        if (_cachedInventoryItemsSaturation != null && (now - _lastInventorySaturationCheck) < 200)
        {
            return _cachedInventoryItemsSaturation;
        }

        _lastInventorySaturationCheck = now;
        _cachedInventoryItemsSaturation = BuildInventoryItemsSaturationMap();
        return _cachedInventoryItemsSaturation;
    }

    private Dictionary<long, bool> BuildInventoryItemsSaturationMap()
    {
        var saturationMap = new Dictionary<long, bool>();

        try
        {
            var inventoryPanel = InGameState.IngameUi.InventoryPanel;
            if (inventoryPanel == null || !inventoryPanel.IsVisible) return saturationMap;

            var inventoryElement = inventoryPanel[InventoryIndex.PlayerInventory];
            if (inventoryElement == null || inventoryElement.Children == null) return saturationMap;

            foreach (var child in inventoryElement.Children)
            {
                if (child == null || !child.IsVisible) continue;

                try
                {
                    var entity = child.Entity;
                    if (entity != null && entity.Address != 0)
                    {
                        // Store saturation state: true if saturated or null, false if explicitly false
                        saturationMap[entity.Address] = child.IsSaturated != false;
                    }
                }
                catch { }
            }
        }
        catch { }

        return saturationMap;
    }

    private void StartBackgroundRetryTask(List<ServerInventory.InventSlotItem> allItems, int maxRetries)
    {
        // Start background task to check and retry missed items
        if (_saturationUpdateTask == null || _saturationUpdateTask.IsCompleted)
        {
            _saturationUpdateTask = Task.Run(async () =>
            {
                try
                {
                    int retryCount = 0;
                    var waitDelay = Settings.FastMode.Value ? 300 : 600;

                    while (retryCount < maxRetries)
                    {
                        // Wait for items to be processed
                        await Task.Delay(waitDelay);

                        // Update saturation map
                        var saturationMap = BuildInventoryItemsSaturationMap();
                        _cachedInventoryItemsSaturation = saturationMap;
                        _lastInventorySaturationCheck = Environment.TickCount64;

                        // Find items that are still in inventory and unsaturated
                        var currentItems = GameController.IngameState.ServerData.PlayerInventories[0].Inventory.InventorySlotItems
                            .Where(x => !IsInIgnoreCell(x))
                            .ToList();

                        var missedItems = new List<ServerInventory.InventSlotItem>();

                        foreach (var originalItem in allItems)
                        {
                            // Check if item still exists in inventory
                            var stillExists = currentItems.Any(x =>
                                x.PosX == originalItem.PosX &&
                                x.PosY == originalItem.PosY &&
                                x.Item?.Address == originalItem.Item?.Address);

                            if (stillExists && originalItem.Item != null && originalItem.Item.Address != 0)
                            {
                                // Check if it's unsaturated
                                if (saturationMap.TryGetValue(originalItem.Item.Address, out var isSaturated) && !isSaturated)
                                {
                                    DebugWindow.LogMsg($"HighlightedItems: Background retry - skipping unsaturated item at ({originalItem.PosX}, {originalItem.PosY})");
                                }
                                else
                                {
                                    // Item is still there and saturated - it was missed
                                    missedItems.Add(originalItem);
                                }
                            }
                        }

                        if (missedItems.Count == 0)
                        {
                            DebugWindow.LogMsg($"HighlightedItems: Background check - no missed items, stopping");
                            break;
                        }

                        DebugWindow.LogMsg($"HighlightedItems: Background retry {retryCount + 1}/{maxRetries} - found {missedItems.Count} missed items");

                        // Click missed items in background
                        await ClickItemsInBackground(missedItems);

                        retryCount++;
                    }

                    DebugWindow.LogMsg("HighlightedItems: Background retry task completed");
                }
                catch (Exception ex)
                {
                    DebugWindow.LogMsg($"HighlightedItems: Background retry error: {ex.Message}");
                }
            });
        }
    }

    private async Task ClickItemsInBackground(List<ServerInventory.InventSlotItem> items)
    {
        try
        {
            // Check if panels are still open
            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                DebugWindow.LogMsg("HighlightedItems: Background click - inventory closed");
                return;
            }

            if (!IsStashTargetOpened)
            {
                DebugWindow.LogMsg("HighlightedItems: Background click - target closed");
                return;
            }

            // Save initial mouse position
            var savedMousePos = Mouse.GetCursorPosition();

            // Press Ctrl before clicking
            Keyboard.KeyDown(Keys.LControlKey);
            await Task.Delay(Settings.FastMode.Value ? 5 : 10);

            // Click each missed item (slower than main loop for better reliability)
            foreach (var item in items)
            {
                try
                {
                    var rect = item.GetClientRect();
                    var position = rect.Center + WindowOffset;

                    Mouse.moveMouse(position);
                    await Task.Delay(Settings.FastMode.Value ? 10 : 30);

                    Mouse.LeftDown();
                    await Task.Delay(Settings.FastMode.Value ? 15 : 35 + Settings.ExtraDelay.Value);

                    Mouse.LeftUp();
                    await Task.Delay(Settings.FastMode.Value ? 5 : 10);
                }
                catch (Exception ex)
                {
                    DebugWindow.LogMsg($"HighlightedItems: Background click error: {ex.Message}");
                }
            }

            // Release Ctrl after clicking
            Keyboard.KeyUp(Keys.LControlKey);
            await Task.Delay(Settings.FastMode.Value ? 5 : 10);

            // Restore mouse position
            Mouse.moveMouse(savedMousePos);
            DebugWindow.LogMsg("HighlightedItems: Background click completed, mouse restored");
        }
        catch (Exception ex)
        {
            DebugWindow.LogMsg($"HighlightedItems: ClickItemsInBackground error: {ex.Message}");
        }
    }

    private async SyncTask<bool> MoveItemsToStash(List<ServerInventory.InventSlotItem> items)
    {
        if (!await MoveItemsCommonPreamble())
        {
            return false;
        }

        // Save initial mouse position and press Ctrl ONCE at the start
        _prevMousePos = Mouse.GetCursorPosition();
        Keyboard.KeyDown(Keys.LControlKey);
        await Wait(KeyDelay, true);

        for (var i = 0; i < items.Count; i++)
        {
            var item = items[i];
            _itemsToMove = items[i..].Select(x => x.GetClientRect()).ToList();

            if (MoveCancellationRequested)
            {
                DebugWindow.LogMsg("HighlightedItems: Cancellation requested");
                await StopMovingItems();
                return false;
            }

            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                DebugWindow.LogMsg("HighlightedItems: Inventory Panel closed, aborting loop");
                await StopMovingItems();
                return false;
            }

            if (!IsStashTargetOpened)
            {
                DebugWindow.LogMsg("HighlightedItems: Target inventory closed, aborting loop");
                await StopMovingItems();
                return false;
            }

            // Add extra delay before first click to ensure Ctrl is registered
            if (i == 0)
            {
                await Wait(TimeSpan.FromMilliseconds(30), true);
            }

            await MoveItem(item.GetClientRect().Center);
        }

        await StopMovingItems();
        return true;
    }

    private bool IsStashTargetOpened =>
        !Settings.VerifyTargetInventoryIsOpened
        || InGameState.IngameUi.StashElement.IsVisible
        || InGameState.IngameUi.SellWindow.IsVisible
        || InGameState.IngameUi.TradeWindow.IsVisible
        || InGameState.IngameUi.GuildStashElement.IsVisible;

    private bool IsStashSourceOpened =>
        !Settings.VerifyTargetInventoryIsOpened
        || InGameState.IngameUi.StashElement.IsVisible
        || InGameState.IngameUi.GuildStashElement.IsVisible;

    private List<RectangleF> _itemsToMove = null;
    private Point _prevMousePos = Point.Zero;

    private async SyncTask<bool> MoveItemsToInventoryWithRetry(Inventory inventory, Predicate<NormalInventoryItem> itemFilter)
    {
        int retryAttempt = 0;
        int maxRetries = Settings.RetryMissedItems ? Settings.MaxRetryAttempts.Value : 0;

        while (retryAttempt <= maxRetries)
        {
            // Get current highlighted items
            var highlightedItems = GetHighlightedItems(inventory, itemFilter)
                .OrderBy(stashItem => stashItem.GetClientRectCache.X)
                .ThenBy(stashItem => stashItem.GetClientRectCache.Y)
                .ToList();

            if (highlightedItems.Count == 0)
            {
                if (retryAttempt == 0)
                {
                    DebugWindow.LogMsg("HighlightedItems: No items to move");
                }
                else
                {
                    DebugWindow.LogMsg("HighlightedItems: All items moved successfully");
                }
                return true;
            }

            if (retryAttempt > 0)
            {
                DebugWindow.LogMsg($"HighlightedItems: Retry attempt {retryAttempt}/{maxRetries}, {highlightedItems.Count} items remaining");
            }

            // Record initial count before moving
            int initialCount = (int)GameController.IngameState.ServerData.PlayerInventories[0].Inventory.CountItems;

            // Move items
            bool success = await MoveItemsToInventory(highlightedItems);

            if (!success)
            {
                return false;
            }

            // Wait for inventory to actually change (items transferred)
            await WaitForInventoryChange(initialCount, 2000);

            // Check if we should retry
            if (!Settings.RetryMissedItems || retryAttempt >= maxRetries)
            {
                break;
            }

            retryAttempt++;
        }

        return true;
    }

    private async SyncTask<bool> MoveItemsToInventory(List<NormalInventoryItem> items)
    {
        if (!await MoveItemsCommonPreamble())
        {
            return false;
        }

        // Calculate how many items we can move based on CountItems limit (max 60)
        var inventory = GameController.IngameState.ServerData.PlayerInventories[0].Inventory;
        int currentCount = (int)inventory.CountItems;
        int maxItems = 60;
        int freeSlots = maxItems - currentCount;

        if (freeSlots <= 0)
        {
            DebugWindow.LogMsg($"HighlightedItems: Inventory full ({currentCount}/{maxItems}), cannot move items");
            return false;
        }

        // Limit items to move to the number of free slots
        int itemsToMove = Math.Min(items.Count, freeSlots);
        DebugWindow.LogMsg($"HighlightedItems: Moving {itemsToMove} items (free slots: {freeSlots}, current: {currentCount}/{maxItems})");

        // Save initial mouse position and press Ctrl ONCE at the start
        _prevMousePos = Mouse.GetCursorPosition();
        Keyboard.KeyDown(Keys.LControlKey);
        await Wait(KeyDelay, true);

        for (var i = 0; i < itemsToMove; i++)
        {
            var item = items[i];
            _itemsToMove = items[i..itemsToMove].Select(x => x.GetClientRectCache).ToList();

            if (MoveCancellationRequested)
            {
                DebugWindow.LogMsg("HighlightedItems: Cancellation requested");
                await StopMovingItems();
                return false;
            }

            if (!IsStashSourceOpened)
            {
                DebugWindow.LogMsg("HighlightedItems: Stash Panel closed, aborting loop");
                await StopMovingItems();
                return false;
            }

            if (!InGameState.IngameUi.InventoryPanel.IsVisible)
            {
                DebugWindow.LogMsg("HighlightedItems: Inventory Panel closed, aborting loop");
                await StopMovingItems();
                return false;
            }

            // Add extra delay before first click to ensure Ctrl is registered
            if (i == 0)
            {
                await Wait(TimeSpan.FromMilliseconds(30), true);
            }

            await MoveItem(item.GetClientRect().Center);
        }

        await StopMovingItems();
        return true;
    }

    private async SyncTask<bool> StopMovingItems() {
        Keyboard.KeyUp(Keys.LControlKey);
        await Wait(KeyDelay, false);
        Mouse.moveMouse(_prevMousePos);
        _prevMousePos = Point.Zero;
        _itemsToMove = null;
        DebugWindow.LogMsg("HighlightedItems: Stopped moving items");
        return true;
    }

    private List<NormalInventoryItem> GetHighlightedItems(Inventory stash, Predicate<NormalInventoryItem> filter)
    {
        try
        {
            var stashItems = stash.VisibleInventoryItems;

            var highlightedItems = stashItems
                .Where(stashItem => filter(stashItem))
                .Where(stashItem => stashItem?.IsSaturated != false) // Only saturated items (skip false, allow null/true)
                .ToList();

            return highlightedItems;
        }
        catch
        {
            return [];
        }
    }

    private TimeSpan KeyDelay => Settings.FastMode ? TimeSpan.FromMilliseconds(5) : TimeSpan.FromMilliseconds(10);
    private TimeSpan MouseMoveDelay => Settings.FastMode ? TimeSpan.FromMilliseconds(5) : TimeSpan.FromMilliseconds(20);
    private TimeSpan MouseDownDelay => Settings.FastMode
        ? TimeSpan.FromMilliseconds(10)
        : TimeSpan.FromMilliseconds(25 + Settings.ExtraDelay.Value);
    private TimeSpan MouseUpDelay => Settings.FastMode ? TimeSpan.FromMilliseconds(2) : TimeSpan.FromMilliseconds(5);
    private int EffectiveIdleMouseDelay => Settings.FastMode ? 50 : Settings.IdleMouseDelay.Value;

    private async SyncTask<bool> MoveItem(SharpDX.Vector2 itemPosition)
    {
        itemPosition += WindowOffset;
        Mouse.moveMouse(itemPosition);
        await Wait(MouseMoveDelay, true);
        Mouse.LeftDown();
        await Wait(MouseDownDelay, true);
        Mouse.LeftUp();
        await Wait(MouseUpDelay, true);
        return true;
    }

    private async SyncTask<bool> Wait(TimeSpan period, bool canUseThreadSleep)
    {
        if (canUseThreadSleep && Settings.UseThreadSleep)
        {
            Thread.Sleep(period);
            return true;
        }

        var sw = Stopwatch.StartNew();
        while (sw.Elapsed < period)
        {
            await TaskUtils.NextFrame();
        }

        return true;
    }

    private readonly ConcurrentDictionary<RectangleF, bool?> _mouseStateForRect = [];

    private bool IsButtonPressed(RectangleF buttonRect)
    {
        var prevState = _mouseStateForRect.GetValueOrDefault(buttonRect);
        var isHovered = buttonRect.Contains(Mouse.GetCursorPosition() - WindowOffset);
        if (!isHovered)
        {
            _mouseStateForRect[buttonRect] = null;
            return false;
        }

        var isPressed = Control.MouseButtons == MouseButtons.Left && CanClickButtons;
        _mouseStateForRect[buttonRect] = isPressed;
        return isPressed &&
               prevState == false;
    }

    private bool CanClickButtons => !Settings.VerifyButtonIsNotObstructed || !ImGui.GetIO().WantCaptureMouse;

    private bool IsInIgnoreCell(ServerInventory.InventSlotItem inventItem)
    {
        var inventPosX = inventItem.PosX;
        var inventPosY = inventItem.PosY;

        if (inventPosX < 0 || inventPosX >= 12)
            return true;
        if (inventPosY < 0 || inventPosY >= 5)
            return true;

        return Settings.IgnoredCells[inventPosY, inventPosX]; //No need to check all item size
    }

    private void DrawIgnoredCellsSettings()
    {
        ImGui.BeginChild("##IgnoredCellsMain", new Vector2(ImGui.GetContentRegionAvail().X, 204f), ImGuiChildFlags.Border,
            ImGuiWindowFlags.NoScrollWithMouse);
        ImGui.Text("Ignored Inventory Slots (checked = ignored)");

        var contentRegionAvail = ImGui.GetContentRegionAvail();
        ImGui.BeginChild("##IgnoredCellsCels", new Vector2(contentRegionAvail.X, contentRegionAvail.Y), ImGuiChildFlags.Border,
            ImGuiWindowFlags.NoScrollWithMouse);

        for (int y = 0; y < 5; ++y)
        {
            for (int x = 0; x < 12; ++x)
            {
                bool isCellIgnored = Settings.IgnoredCells[y, x];
                if (ImGui.Checkbox($"##{y}_{x}IgnoredCells", ref isCellIgnored))
                    Settings.IgnoredCells[y, x] = isCellIgnored;
                if (x < 11)
                    ImGui.SameLine();
            }
        }

        ImGui.EndChild();
        ImGui.EndChild();
    }
}
