import { Dictionary } from './dictionary';
export declare enum StorageType {
    LocalStorage = 0,
    SessionStorage = 1,
}
export interface Listener {
    subscribe(): Subscription;
    subscribe(next?: () => void, error?: (error: any) => void, complete?: () => void): Subscription;
}
export interface Subscription {
    /**
     * A flag to indicate whether this Subscription has already been unsubscribed.
     * @type {boolean}
     */
    closed: boolean;
    /**
     * Disposes the resources held by the subscription. May, for instance, cancel
     * an ongoing Observable execution or cancel any other type of work that
     * started when the Subscription was created.
     * @return {void}
     */
    unsubscribe(): void;
}
/**
 * Helper for creating and querying Local Storage or Session Storage.
 * Uses {@link Dictionary} so all the data is encapsulated in a single
 * storage namespace. Writes update the actual storage.
 */
export declare class Storage<T> extends Dictionary<T> {
    container: string;
    private _type;
    private _storage;
    private readonly _current;
    /**
     * @constructor
     * @param {string} container Container name to be created in the LocalStorage.
     * @param {StorageType} type[optional] Storage Type to be used, defaults to Local Storage.
    */
    constructor(container: string, _type?: StorageType);
    /**
     * Switch the storage type.
     * Switches the storage type and then reloads the in-memory collection.
     *
     * @type {StorageType} type The desired storage to be used.
     */
    switchStorage(type: StorageType): void;
    /**
     * Add an item.
     * Extends Dictionary's implementation of add, with a save to the storage.
     */
    add(item: string, value: T): T;
    /**
     * Add or Update an item.
     * Extends Dictionary's implementation of insert, with a save to the storage.
     */
    insert(item: string, value: T): T;
    /**
     * Remove an item.
     * Extends Dictionary's implementation with a save to the storage.
     */
    remove(item: string): T;
    /**
     * Clear the storage.
     * Extends Dictionary's implementation with a save to the storage.
     */
    clear(): void;
    /**
     * Clear all storages.
     * Completely clears both the localStorage and sessionStorage.
     */
    static clearAll(): void;
    /**
     * Refreshes the storage with the current localStorage values.
     */
    load(): void;
    /**
     * Notify that the storage has changed only if the 'notify'
     * property has been subscribed to.
     */
    notify: () => Listener;
    /**
     * Synchronizes the current state to the storage.
     */
    private _sync(item, value);
}
