export class PagedLoader<T> {
    private iterator: AsyncGenerator<T[], void, unknown>;
    private done = false;

    public items: T[] = [];

    constructor(iterator: AsyncGenerator<T[], void, unknown>) {
        this.iterator = iterator;
    }

    public get hasMore(): boolean {
        return !this.done;
    }

    public async loadNextPage(): Promise<T[]> {
        if (this.done) return [];

        const result = await this.iterator.next();

        if (result.done) {
            this.done = true;
            return [];
        }

        const page = result.value;
        this.items.push(...page);
        return page;
    }

    public reset(iterator: AsyncGenerator<T[], void, unknown>): void {
        this.iterator = iterator;
        this.items = [];
        this.done = false;
    }
}
