using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace DocHandler.Helpers
{
    /// <summary>
    /// Extension methods for ObservableCollection to improve performance and thread safety
    /// </summary>
    public static class ObservableCollectionExtensions
    {
        /// <summary>
        /// Efficiently updates an ObservableCollection from a source on the UI thread
        /// </summary>
        public static async Task UpdateFromSourceAsync<T>(
            this ObservableCollection<T> collection,
            IEnumerable<T> source,
            IEqualityComparer<T>? comparer = null,
            DispatcherPriority priority = DispatcherPriority.DataBind,
            CancellationToken cancellationToken = default)
        {
            if (collection == null) throw new ArgumentNullException(nameof(collection));
            if (source == null) throw new ArgumentNullException(nameof(source));
            
            var sourceList = source as IList<T> ?? source.ToList();
            comparer ??= EqualityComparer<T>.Default;
            
            // Skip update if collections are equal
            if (collection.SequenceEqual(sourceList, comparer))
                return;
            
            // Ensure we're on the UI thread
            if (Application.Current?.Dispatcher == null)
                throw new InvalidOperationException("No dispatcher available");
            
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                collection.Clear();
                foreach (var item in sourceList)
                {
                    if (cancellationToken.IsCancellationRequested)
                        break;
                    collection.Add(item);
                }
            }, priority);
        }
        
        /// <summary>
        /// Safely updates collection with null and disposal checks
        /// </summary>
        public static async Task SafeUpdateAsync<T>(
            this ObservableCollection<T> collection,
            Func<Task<IEnumerable<T>>> sourceFactory,
            Action<Exception>? onError = null,
            CancellationToken cancellationToken = default)
        {
            try
            {
                var source = await sourceFactory().ConfigureAwait(false);
                
                if (!cancellationToken.IsCancellationRequested && 
                    Application.Current?.Dispatcher != null &&
                    !Application.Current.Dispatcher.HasShutdownStarted)
                {
                    await collection.UpdateFromSourceAsync(source, cancellationToken: cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
                // Expected, don't report
            }
            catch (Exception ex)
            {
                onError?.Invoke(ex);
            }
        }
        
        /// <summary>
        /// Adds a range of items efficiently
        /// </summary>
        public static async Task AddRangeAsync<T>(
            this ObservableCollection<T> collection,
            IEnumerable<T> items,
            DispatcherPriority priority = DispatcherPriority.DataBind,
            CancellationToken cancellationToken = default)
        {
            if (collection == null) throw new ArgumentNullException(nameof(collection));
            if (items == null) throw new ArgumentNullException(nameof(items));
            
            var itemsList = items as IList<T> ?? items.ToList();
            if (!itemsList.Any()) return;
            
            if (Application.Current?.Dispatcher == null)
                throw new InvalidOperationException("No dispatcher available");
            
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                foreach (var item in itemsList)
                {
                    if (cancellationToken.IsCancellationRequested)
                        break;
                    collection.Add(item);
                }
            }, priority);
        }
        
        /// <summary>
        /// Replaces all items efficiently without clearing first
        /// </summary>
        public static async Task ReplaceAllAsync<T>(
            this ObservableCollection<T> collection,
            IEnumerable<T> newItems,
            DispatcherPriority priority = DispatcherPriority.DataBind,
            CancellationToken cancellationToken = default)
        {
            if (collection == null) throw new ArgumentNullException(nameof(collection));
            if (newItems == null) throw new ArgumentNullException(nameof(newItems));
            
            var newItemsList = newItems as IList<T> ?? newItems.ToList();
            
            if (Application.Current?.Dispatcher == null)
                throw new InvalidOperationException("No dispatcher available");
            
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                // Suspend notifications for batch update
                var collectionView = System.Windows.Data.CollectionViewSource.GetDefaultView(collection);
                using (collectionView?.DeferRefresh())
                {
                    collection.Clear();
                    foreach (var item in newItemsList)
                    {
                        if (cancellationToken.IsCancellationRequested)
                            break;
                        collection.Add(item);
                    }
                }
            }, priority);
        }
    }
} 