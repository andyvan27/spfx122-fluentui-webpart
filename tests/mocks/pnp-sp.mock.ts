// A minimal interface describing the fluent items chain
interface IItemsChain {
  select: jest.Mock<IItemsChain, any>;
  expand: jest.Mock<() => any, any>;
}

// Create ONE shared mock instance
const finalCall = jest.fn();
const chain: IItemsChain = {
  select: jest.fn(() => chain),
  expand: jest.fn(() => finalCall)
};

const spInstance = {
  using: jest.fn().mockReturnThis(),
  web: {
    lists: {
      getByTitle: jest.fn((title: string) => ({
        items: chain
      }))
    }
  }
};

export const spfi = jest.fn(() => spInstance);
export const SPFx = jest.fn();

// Export chain + finalCall so tests can override them
export const __chain = chain;
export const __finalCall = finalCall;
