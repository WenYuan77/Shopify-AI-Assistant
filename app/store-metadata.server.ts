type GraphQLAdmin = { graphql: (query: string, options?: object) => Promise<Response> };

export interface StoreMetadata {
  /** 是否有产品 */
  hasProducts: boolean;
  /** 是否有订单 */
  hasOrders: boolean;
  /** 订单日期范围 */
  ordersDateRange: { from: string; to: string } | null;
  /** 产品数量 */
  productsCount: number;
  /** 订单数量 */
  ordersCount: number;
  /** 客户数量（需 read_customers 权限） */
  customersCount: number | null;
  /** 产品系列数量 */
  collectionsCount: number | null;
}

export async function getStoreMetadata(admin: GraphQLAdmin): Promise<StoreMetadata> {
  const results = await Promise.allSettled([
    admin.graphql(`#graphql query { products(first: 1) { edges { node { id } } } }`),
    admin.graphql(`#graphql
      query { orders(first: 1, sortKey: CREATED_AT, reverse: false, query: "status:any") { edges { node { createdAt } } } }
    `),
    admin.graphql(`#graphql
      query { orders(first: 1, sortKey: CREATED_AT, reverse: true, query: "status:any") { edges { node { createdAt } } } }
    `),
    admin.graphql(`#graphql query { productsCount { count } }`),
    admin.graphql(`#graphql query { ordersCount { count } }`),
    admin.graphql(`#graphql query { customersCount { count } }`),
    admin.graphql(`#graphql query { collectionsCount { count } }`),
  ]);

  const [productsJson, ordersFirstJson, ordersLastJson, productsCountRes, ordersCountRes, customersCountRes, collectionsCountRes] =
    await Promise.all(
      results.map((r) => (r.status === "fulfilled" ? r.value.json() : Promise.resolve(null)))
    );

  const hasProducts = ((productsJson as { data?: { products?: { edges?: unknown[] } } })?.data?.products?.edges?.length ?? 0) > 0;
  const firstOrder = (ordersFirstJson as { data?: { orders?: { edges?: { node?: { createdAt?: string } }[] } } })?.data?.orders?.edges?.[0]?.node;
  const lastOrder = (ordersLastJson as { data?: { orders?: { edges?: { node?: { createdAt?: string } }[] } } })?.data?.orders?.edges?.[0]?.node;
  const hasOrders = Boolean(firstOrder || lastOrder);

  let ordersDateRange: { from: string; to: string } | null = null;
  if (firstOrder?.createdAt && lastOrder?.createdAt) {
    const from = new Date(firstOrder.createdAt);
    const to = new Date(lastOrder.createdAt);
    ordersDateRange = {
      from: `${from.getFullYear()}-${String(from.getMonth() + 1).padStart(2, "0")}`,
      to: `${to.getFullYear()}-${String(to.getMonth() + 1).padStart(2, "0")}`,
    };
  } else if (firstOrder?.createdAt) {
    const d = new Date(firstOrder.createdAt);
    const str = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    ordersDateRange = { from: str, to: str };
  } else if (lastOrder?.createdAt) {
    const d = new Date(lastOrder.createdAt);
    const str = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    ordersDateRange = { from: str, to: str };
  }

  const productsCount = (productsCountRes as { data?: { productsCount?: { count?: number } } })?.data?.productsCount?.count ?? 0;
  const ordersCount = (ordersCountRes as { data?: { ordersCount?: { count?: number } } })?.data?.ordersCount?.count ?? 0;
  let customersCount: number | null = null;
  if (customersCountRes && (customersCountRes as { data?: { customersCount?: { count?: number } } })?.data?.customersCount) {
    customersCount = (customersCountRes as { data: { customersCount: { count: number } } }).data.customersCount.count;
  }

  let collectionsCount: number | null = null;
  if (collectionsCountRes && (collectionsCountRes as { data?: { collectionsCount?: { count?: number } } })?.data?.collectionsCount) {
    collectionsCount = (collectionsCountRes as { data: { collectionsCount: { count: number } } }).data.collectionsCount.count;
  }

  return {
    hasProducts,
    hasOrders,
    ordersDateRange,
    productsCount,
    ordersCount,
    customersCount,
    collectionsCount,
  };
}
