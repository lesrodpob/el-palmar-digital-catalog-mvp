import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  Home, Wine, Martini, Beer, CupSoda, Snowflake,
  Search, ShoppingCart, MessageCircle, Trash2,
  Minus, Plus, Package, Truck, Percent, Headphones, ChevronRight,
  LoaderCircle, AlertCircle
, Banknote, FileText, CreditCard, ArrowLeftRight} from "lucide-react";
import "./styles.css";

const EXCEL_URL = "/data/productos.xlsx";

const categoryGroups = [
  { key: "ALL", label: "Inicio", icon: Home },
  { key: "VINOS", label: "Vinos", icon: Wine },
  { key: "ESPUMANTES", label: "Espumantes", icon: Wine },
  { key: "LICORES", label: "Licores", icon: Martini },
  { key: "CERVEZAS", label: "Cervezas", icon: Beer },
  { key: "BEBIDAS", label: "Bebidas", icon: CupSoda },
  { key: "CONGELADOS", label: "Congelados", icon: Snowflake },
];

const groupMap = {
  VINOS: ["VINOS"],
  ESPUMANTES: ["ESPUMANTES"],
  LICORES: ["LICORES"],
  CERVEZAS: ["CERVEZAS"],
  BEBIDAS: ["BEBIDAS NO ALCOHOLICAS"],
  CONGELADOS: ["CONGELADOS"],
};

const money = (value) => new Intl.NumberFormat("es-CL", {
  style: "currency", currency: "CLP", maximumFractionDigits: 0
}).format(Number(value) || 0);

function normalizeProduct(row, index) {
  const stock = Number(row.stock);
  return {
    id: String(row.id ?? index),
    barcode: String(row.barcode ?? "").replace(/\.0$/, ""),
    category: String(row.category ?? "").trim(),
    family: String(row.family ?? "").trim(),
    name: String(row.name ?? "").trim(),
    price: Number(row.price) || 0,
    stock: Number.isFinite(stock) ? stock : 0,
  };
}

async function loadProductsFromExcel() {
  const response = await fetch(`${EXCEL_URL}?v=${Date.now()}`);
  if (!response.ok) throw new Error(`No se pudo cargar ${EXCEL_URL}`);
  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" });
  return rows
    .map((row, index) => normalizeProduct({
      id: row.ID,
      barcode: row["Código Barra"],
      category: row.Categoría,
      family: row.Familia,
      name: row["Nombre Producto"],
      price: row.Precio,
      stock: row.MATRIZ,
    }, index))
    .filter((p) => p.name && p.category);
}

function App() {
  const [products, setProducts] = useState([]);
  const [selectedCategory, setSelectedCategory] = useState("ALL");
  const [selectedFamily, setSelectedFamily] = useState("ALL");
  const [search, setSearch] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [cart, setCart] = useState([]);
  const [cartOpen, setCartOpen] = useState(false);
  const [infoOpen, setInfoOpen] = useState(true);
  const cartRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    setCurrentPage(1);
    setSelectedFamily("ALL");
  }, [search]);

  useEffect(() => {
    const query = search.trim();

    if (!query) return;

    const timer = setTimeout(() => {
      document.getElementById("products-section")?.scrollIntoView({
        behavior: "smooth",
        block: "start"
      });
    }, 350);

    return () => clearTimeout(timer);
  }, [search]);

  useEffect(() => {
    if (cartOpen) {
      requestAnimationFrame(() => {
        cartRef.current?.scrollTo({
          top: 0,
          behavior: "auto"
        });
      });
    }
  }, [cartOpen]);

  useEffect(() => {
    loadProductsFromExcel()
      .then(setProducts)
      .catch((err) => setError(err.message))
      .finally(() => setLoading(false));
  }, []);

  const featuredProducts = useMemo(
    () => products.filter(p => p.category === "DESTACADOS"),
    [products]
  );

  const searchCategory = useMemo(() => {
    const query = search.toLowerCase().trim();

    if (!query) return null;

    return categoryGroups.find(
      ({ key, label }) =>
        key !== "ALL" &&
        label.toLowerCase() === query
    )?.key || null;
  }, [search]);

  const availableFamilies = useMemo(() => {
    const categoryForFamilies = searchCategory || selectedCategory;

    const cats = categoryForFamilies === "ALL"
      ? Object.values(groupMap).flat()
      : groupMap[categoryForFamilies];

    const families = products
      .filter(p => cats?.includes(p.category))
      .map(p => p.family)
      .filter(Boolean);

    return [...new Set(families)].sort((a, b) => a.localeCompare(b, "es"));
  }, [products, selectedCategory, searchCategory]);

  const filteredProducts = useMemo(() => {
    const query = search.toLowerCase().trim();
    const categoryForSearch = searchCategory || selectedCategory;
    const cats = categoryForSearch === "ALL" ? null : groupMap[categoryForSearch];
    return products.filter((p) => {
      const matchesCategory = !cats || cats.includes(p.category);
      const matchesFamily = selectedFamily === "ALL" || p.family === selectedFamily;
      const matchesSearch = !query || [p.name, p.family, p.category, p.barcode]
        .some(v => v.toLowerCase().includes(query));
      return matchesCategory &&
        matchesFamily &&
        matchesSearch &&
        p.category !== "PROMOCIONES" &&
        p.category !== "DESTACADOS";
    });
  }, [products, selectedCategory, selectedFamily, search, searchCategory]);

  const cartUnits = cart.reduce((sum, item) => sum + item.quantity, 0);

  // Descuento de hielo por formato:
  // - Hielo de 2 kg: desde 30 bolsas -> $100 menos por bolsa.
  // - Hielo de 1 kg: desde 50 bolsas -> $100 menos por bolsa.
  // Cada formato se calcula por separado.
  const ice2kgUnits = cart
    .filter(item => {
      const text = `${item.name} ${item.family}`.toLowerCase();
      return /hielo/.test(text) && /\b2\s*(?:kg|kl|kilos?)\b/i.test(text);
    })
    .reduce((sum, item) => sum + item.quantity, 0);

  const ice1kgUnits = cart
    .filter(item => {
      const text = `${item.name} ${item.family}`.toLowerCase();
      return /hielo/.test(text) && /\b1\s*(?:kg|kl|kilos?)\b/i.test(text);
    })
    .reduce((sum, item) => sum + item.quantity, 0);

  const cartSubtotal = cart.reduce(
    (sum, item) => sum + item.price * item.quantity,
    0
  );

  const iceDiscount =
    (ice2kgUnits >= 30 ? ice2kgUnits * 100 : 0) +
    (ice1kgUnits >= 50 ? ice1kgUnits * 100 : 0);

  const cartTotal = cartSubtotal - iceDiscount;

  const PRODUCTS_PER_PAGE = 24;
  const totalPages = Math.ceil(filteredProducts.length / PRODUCTS_PER_PAGE);
  const paginatedProducts = filteredProducts.slice(
    (currentPage - 1) * PRODUCTS_PER_PAGE,
    currentPage * PRODUCTS_PER_PAGE
  );

  function selectCategory(key) {
    setSelectedCategory(key);
    setCurrentPage(1);
    setSelectedFamily("ALL");
    setTimeout(() => document.getElementById("products-section")?.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
  }

  function goHome() {
    setSelectedCategory("ALL");
    setCurrentPage(1);
    setSelectedFamily("ALL");
    setSearch("");
    document.querySelector(".main")?.scrollTo({
      top: 0,
      behavior: "smooth"
    });
  }

  function scrollToProducts() {
    setTimeout(() => {
      document.getElementById("products-section")?.scrollIntoView({
        behavior: "smooth",
        block: "start"
      });
    }, 50);
  }

  function addToCart(product) {
    if (product.stock <= 0) return;

    setCartOpen(true);

    setCart(prev => {
      const existing = prev.find(i => i.id === product.id);
      if (existing) {
        if (existing.quantity >= Math.floor(product.stock)) return prev;
        return prev.map(i => i.id === product.id ? { ...i, quantity: i.quantity + 1 } : i);
      }
      return [...prev, { ...product, quantity: 1 }];
    });
  }

  function changeQty(id, delta) {
    setCart(prev => prev.flatMap(item => {
      if (item.id !== id) return [item];
      const next = item.quantity + delta;
      if (next <= 0) return [];
      if (next > Math.floor(item.stock)) return [item];
      return [{ ...item, quantity: next }];
    }));
  }

  function setQty(id, value) {
    const max = Math.max(1, Math.floor(
      cart.find(item => item.id === id)?.stock || 1
    ));
    const next = Math.min(Math.max(1, Number(value) || 1), max);

    setCart(prev =>
      prev.map(item =>
        item.id === id ? { ...item, quantity: next } : item
      )
    );
  }

  function openWhatsAppContact() {
    const whatsappNumber = "988137633";
    window.open(`https://wa.me/${whatsappNumber}`, "_blank");
  }

  function sendWhatsApp() {
    if (!cart.length) return;

    const whatsappNumber = "988137633";
    const lines = cart.map(
      i => `• ${i.name} x${i.quantity} — ${money(i.price * i.quantity)}`
    );

    const message = `\u{1F3EA} Hola Distribuidora El Palmar

\u{1F6D2} Quisiera cotizar el siguiente pedido:

${lines.join("\n")}

\u{1F4B0} Total: ${money(cartTotal)}

\u{1F4CB} Quedo atento/a para confirmar el pedido.
¡Muchas gracias! \u{1F60A}`;

    window.open(
      `https://wa.me/${whatsappNumber}?text=${encodeURIComponent(message)}`,
      "_blank"
    );
  }

  const hero = (
    <section className="hero">
      <img src="/el-palmar-banner.png" alt="Distribuidora El Palmar" />
    </section>
  );

  return (
    <div className="app">
      <aside className="sidebar">
        <div className="side-title-row">
          <div className="side-title">CATEGORÍAS</div>
          <div className="category-scroll-hint" aria-hidden="true">
            <span>Desliza para más categorías</span>
            <ChevronRight size={16} />
          </div>
        </div>
        <div className="side-line" />
        <nav>
          <button
            className={`side-item ${selectedCategory === "ALL" && !search ? "active" : ""}`}
            onClick={goHome}
          >
            <Home size={21} /><span>Inicio</span>
          </button>

          {categoryGroups
            .filter(({ key }) => key !== "ALL")
            .map(({ key, label, icon: Icon }) => (
              <button
                key={key}
                className={`side-item ${selectedCategory === key ? "active" : ""}`}
                onClick={() => selectCategory(key)}
              >
                <Icon size={21} /><span>{label}</span>
              </button>
            ))}
        </nav>
        <div className="side-tag">Tu distribuidora<br />de confianza</div>
      </aside>

      <main className="main">
        {hero}
        <div className="search-wrap">
          <Search size={27} />
          <input
            value={search}
            onChange={e => setSearch(e.target.value)}
            onKeyDown={e => {
              if (e.key === "Enter" && search.trim()) {
                scrollToProducts();
              }
            }}
            placeholder="Buscar productos, marcas o categorías..."
          />
        </div>

        <div className="info-trigger-wrap">
          <button
            className="info-trigger"
            onClick={() => setInfoOpen(true)}
            type="button"
          >
            <Truck size={18} />
            <span>Despacho y medios de pago</span>
            <ChevronRight size={17} />
          </button>
        </div>

        <section className="content">
          {loading && <div className="state"><LoaderCircle className="spin" /> Cargando catálogo desde Excel...</div>}
          {error && <div className="state error"><AlertCircle /> {error}</div>}

          {!loading && !error && (
            <>
              <section className="promo-section">
                <div className="section-head">
                  <h2>Ofertas y productos destacados</h2>
                </div>

                {products.some(
                  p => p.family.toLowerCase() === "hielos"
                ) && (
                    <div className="ice-deal">
                      <img
                        src="/hielo-volumen-banner.png"
                        alt="Descuento por volumen en hielo"
                      />
                    </div>
                  )}

                <div className="product-grid promo-grid">
                  {featuredProducts.map(p => (
                    <ProductCard key={p.id} product={p} onAdd={addToCart} promo={false} />
                  ))}
                </div>
              </section>

              <section id="products-section" className="products-head">
                <div className="section-head">
                  <h2>{
                    searchCategory
                      ? categoryGroups.find(c => c.key === searchCategory)?.label
                      : selectedCategory === "ALL"
                        ? "Todos los productos"
                        : categoryGroups.find(c => c.key === selectedCategory)?.label
                  }</h2>
                  <span>{filteredProducts.length} productos</span>
                </div>

                <div className="family-filter">
                  <button
                    className={selectedFamily === "ALL" ? "selected" : ""}
                    onClick={() => {
                      setSelectedFamily("ALL");
                      setCurrentPage(1);
                      setTimeout(() => {
                        document.getElementById("products-section")?.scrollIntoView({
                          behavior: "smooth",
                          block: "start"
                        });
                      }, 50);
                    }}
                  >
                    Todas las familias
                  </button>
                  {availableFamilies.map(f => (
                    <button
                      key={f}
                      className={selectedFamily === f ? "selected" : ""}
                      onClick={() => {
                        setSelectedFamily(f);
                        setCurrentPage(1);
                        setTimeout(() => {
                          document.getElementById("products-section")?.scrollIntoView({
                            behavior: "smooth",
                            block: "start"
                          });
                        }, 50);
                      }}
                    >
                      {f}
                    </button>
                  ))}
                </div>

                {totalPages > 1 && (
                  <div className="pagination minimal-pagination">
                    <button
                      disabled={currentPage === 1}
                      onClick={() => {
                        setCurrentPage(page => page - 1);
                        document.getElementById("products-section")?.scrollIntoView({
                          behavior: "smooth",
                          block: "start"
                        });
                      }}
                    >
                      ← Anterior
                    </button>

                    <span className="page-indicator">
                      Página <strong>{currentPage}</strong> de <strong>{totalPages}</strong>
                    </span>

                    <button
                      disabled={currentPage === totalPages}
                      onClick={() => {
                        setCurrentPage(page => page + 1);
                        document.getElementById("products-section")?.scrollIntoView({
                          behavior: "smooth",
                          block: "start"
                        });
                      }}
                    >
                      Siguiente →
                    </button>
                  </div>
                )}

                {filteredProducts.length === 0 ? (
                  <div className="no-results">No encontramos productos con esos filtros.</div>
                ) : (
                  <>
                    <div className="product-grid">
                      {paginatedProducts.map(p => (
                        <ProductCard key={p.id} product={p} onAdd={addToCart} />
                      ))}
                    </div>

                    {totalPages > 1 && (
                      <div className="pagination minimal-pagination">
                        <button
                          disabled={currentPage === 1}
                          onClick={() => {
                            setCurrentPage(page => page - 1);
                            document.getElementById("products-section")?.scrollIntoView({
                              behavior: "smooth",
                              block: "start"
                            });
                          }}
                        >
                          ← Anterior
                        </button>

                        <span className="page-indicator">
                          Página <strong>{currentPage}</strong> de <strong>{totalPages}</strong>
                        </span>

                        <button
                          disabled={currentPage === totalPages}
                          onClick={() => {
                            setCurrentPage(page => page + 1);
                            document.getElementById("products-section")?.scrollIntoView({
                              behavior: "smooth",
                              block: "start"
                            });
                          }}
                        >
                          Siguiente →
                        </button>
                      </div>
                    )}


                  </>
                )}
              </section>
            </>
          )}
        </section>
      </main>

      {!cartOpen && (
        <button
          className="cart-toggle"
          onClick={() => setCartOpen(true)}
          aria-label={`Ver pedido, ${cartUnits} productos`}
        >
          <ShoppingCart size={23} />
          <span>Ver pedido</span>
          <span className="cart-toggle-whatsapp" aria-label="WhatsApp">
            <svg viewBox="0 0 32 32" aria-hidden="true" focusable="false">
              <path d="M16 2.6C8.63 2.6 2.65 8.58 2.65 15.95c0 2.35.62 4.55 1.79 6.49L2.6 29.4l7.13-1.82a13.28 13.28 0 0 0 6.27 1.57h.01c7.36 0 13.34-5.98 13.34-13.34C29.35 8.58 23.37 2.6 16 2.6Z" fill="#25D366"/>
              <path d="M11.07 8.86c.3-.43.75-.64 1.26-.56l1.56.24c.43.07.77.35.9.77l.54 1.73c.11.36.02.75-.25 1.02l-.77.78c-.18.18-.22.45-.09.68.74 1.3 1.77 2.38 3.03 3.19.23.15.52.14.72-.04l.82-.74c.29-.26.71-.34 1.06-.2l1.63.66c.4.16.67.53.68.96l.04 1.54c.01.5-.25.96-.69 1.2-.58.32-1.28.48-1.98.4-1.76-.2-3.71-1.32-5.55-3.16-1.83-1.83-2.96-3.78-3.16-5.54-.08-.71.08-1.41.4-1.99l.85-.94Z" fill="#fff"/>
            </svg>
          </span>
          <b>{cartUnits}</b>
        </button>
      )}

      <aside ref={cartRef} className={`cart ${cartOpen ? "cart-open" : ""}`}>
        <div className="cart-box">
          <div className="cart-head">
            <ShoppingCart size={32} />
            <strong>Tu pedido</strong>
            <span>{cartUnits}</span>
            <button
              className="cart-close"
              onClick={() => setCartOpen(false)}
              aria-label="Cerrar pedido"
            >
              ×
            </button>
          </div>
          <div className="cart-items">
            {cart.length === 0 ? <p>Tu pedido aparecerá aquí.</p> : cart.map(item => (
              <div className="cart-row" key={item.id}>
                <div className="cart-thumb"><Package size={22} /></div>
                <div className="cart-info"><b>{item.name}</b><strong>{money(item.price)}</strong></div>
                <div className="qty">
                  <button onClick={() => changeQty(item.id, -1)} aria-label="Disminuir cantidad"><Minus size={14} /></button>
                  <QuantityInput
                    value={item.quantity}
                    max={Math.floor(item.stock)}
                    onChange={(value) => setQty(item.id, value)}
                  />
                  <button onClick={() => changeQty(item.id, 1)} aria-label="Aumentar cantidad"><Plus size={14} /></button>
                </div>
                <button className="remove" onClick={() => changeQty(item.id, -item.quantity)}><Trash2 size={18} /></button>
              </div>
            ))}
          </div>
          <div className="cart-total">
            <div>
              <span>Total ({cartUnits} productos)</span>
              {iceDiscount > 0 && (
                <small>
                  Descuento hielo: -{money(iceDiscount)}
                </small>
              )}
            </div>
            <b>{money(cartTotal)}</b>
          </div>
          <button className="whatsapp" onClick={sendWhatsApp}>
            <span className="order-whatsapp-icon" aria-hidden="true">
              <svg viewBox="0 0 32 32">
                <path
                  d="M16 2.6C8.63 2.6 2.65 8.58 2.65 15.95c0 2.35.62 4.55 1.79 6.49L2.6 29.4l7.13-1.82a13.28 13.28 0 0 0 6.27 1.57h.01c7.36 0 13.34-5.98 13.34-13.34C29.35 8.58 23.37 2.6 16 2.6Z"
                  fill="#25D366"
                />
                <path
                  d="M11.07 8.86c.3-.43.75-.64 1.26-.56l1.56.24c.43.07.77.35.9.77l.54 1.73c.11.36.02.75-.25 1.02l-.77.78c-.18.18-.22.45-.09.68.74 1.3 1.77 2.38 3.03 3.19.23.15.52.14.72-.04l.82-.74c.29-.26.71-.34 1.06-.2l1.63.66c.4.16.67.53.68.96l.04 1.54c.01.5-.25.96-.69 1.2-.58.32-1.28.48-1.98.4-1.76-.2-3.71-1.32-5.55-3.16-1.83-1.83-2.96-3.78-3.16-5.54-.08-.71.08-1.41.4-1.99l.85-.94Z"
                  fill="#fff"
                />
              </svg>
            </span>
            <span>Enviar pedido por WhatsApp</span>
          </button>
          <button className="clear" onClick={() => setCart([])}><Trash2 size={20} /> Vaciar carrito</button>
        </div>
      </aside>

      {infoOpen && (
        <div className="info-modal-backdrop" onClick={() => setInfoOpen(false)}>
          <div
            className="info-modal info-modal-final"
            role="dialog"
            aria-modal="true"
            aria-labelledby="info-modal-title"
            onClick={e => e.stopPropagation()}
          >
            <div className="info-modal-banner">
              <div className="info-brand-lockup">
                <img
                  src="/el-palmar-modal-banner.png"
                  alt="Distribuidora El Palmar"
                />
              </div>

              <div className="info-brand-phrase">
                <span>Todo lo que necesitas</span>
                <span>en un solo lugar</span>
              </div>

              <button
                className="info-modal-close"
                onClick={() => setInfoOpen(false)}
                aria-label="Cerrar información"
                type="button"
              >
                ×
              </button>

              <div className="info-modal-title">
                <span>INFORMACIÓN PARA TU PEDIDO</span>
                <h3 id="info-modal-title">Despacho y medios de pago</h3>
              </div>
            </div>

            <div className="info-modal-content">
              <section className="info-card info-delivery-card">
                <div className="info-icon">
                  <Truck size={24} />
                </div>

                <div className="info-card-copy">
                  <span className="info-label">DESPACHOS</span>
                  <h4>Despachos gratis por compras desde $50.000</h4>

                  <div className="delivery-list">
                    <p><strong>● Limache:</strong> Martes</p>
                    <p><strong>● Olmué:</strong> Miércoles</p>
                    <p><strong>● Otros sectores:</strong> Comunicarse al WhatsApp.</p>
                  </div>
                </div>
              </section>

              <section className="info-card info-payment-card">
                <div className="info-icon">
                  <CreditCardIcon />
                </div>

                <div className="info-card-copy">
                  <span className="info-label">MEDIOS DE PAGO</span>
                  <h4>Aceptamos los siguientes medios de pago:</h4>

                  <div className="payment-logo-grid">
                    <div className="payment-logo-item">
                      <span className="payment-logo"><Banknote size={19} /></span>
                      <span>Efectivo</span>
                    </div>

                    <div className="payment-logo-item">
                      <span className="payment-logo"><FileText size={19} /></span>
                      <span>Cheque</span>
                    </div>

                    <div className="payment-logo-item">
                      <span className="payment-logo"><CreditCard size={19} /></span>
                      <span>Tarjetas débito</span>
                    </div>

                    <div className="payment-logo-item">
                      <span className="payment-logo"><CreditCard size={19} /></span>
                      <span>Tarjetas crédito</span>
                    </div>

                    <div className="payment-logo-item">
                      <span className="payment-logo"><ArrowLeftRight size={19} /></span>
                      <span>Transferencias</span>
                    </div>
                  </div>
                </div>
              </section>

              <div className="info-modal-footer">
                <div className="info-footer-icon" aria-hidden="true">
                  <svg viewBox="0 0 32 32" aria-hidden="true">
                        <path
                          d="M16 2.6C8.63 2.6 2.65 8.58 2.65 15.95c0 2.35.62 4.55 1.79 6.49L2.6 29.4l7.13-1.82a13.28 13.28 0 0 0 6.27 1.57h.01c7.36 0 13.34-5.98 13.34-13.34C29.35 8.58 23.37 2.6 16 2.6Z"
                          fill="#25D366"
                        />
                        <path
                          d="M11.07 8.86c.3-.43.75-.64 1.26-.56l1.56.24c.43.07.77.35.9.77l.54 1.73c.11.36.02.75-.25 1.02l-.77.78c-.18.18-.22.45-.09.68.74 1.3 1.77 2.38 3.03 3.19.23.15.52.14.72-.04l.82-.74c.29-.26.71-.34 1.06-.2l1.63.66c.4.16.67.53.68.96l.04 1.54c.01.5-.25.96-.69 1.2-.58.32-1.28.48-1.98.4-1.76-.2-3.71-1.32-5.55-3.16-1.83-1.83-2.96-3.78-3.16-5.54-.08-.71.08-1.41.4-1.99l.85-.94Z"
                          fill="#fff"
                        />
                      </svg>
                </div>

                <div className="info-footer-copy">
                  <strong>¿Tienes dudas o consultas?</strong>
                  <span>Escríbenos por WhatsApp y te ayudaremos.</span>
                </div>

                <div className="info-modal-actions">
                  <button type="button" className="info-whatsapp-button" onClick={openWhatsAppContact}>
                    <span className="whatsapp-button-icon" aria-hidden="true">
                      <svg viewBox="0 0 32 32" aria-hidden="true">
                        <path
                          d="M16 2.6C8.63 2.6 2.65 8.58 2.65 15.95c0 2.35.62 4.55 1.79 6.49L2.6 29.4l7.13-1.82a13.28 13.28 0 0 0 6.27 1.57h.01c7.36 0 13.34-5.98 13.34-13.34C29.35 8.58 23.37 2.6 16 2.6Z"
                          fill="#25D366"
                        />
                        <path
                          d="M11.07 8.86c.3-.43.75-.64 1.26-.56l1.56.24c.43.07.77.35.9.77l.54 1.73c.11.36.02.75-.25 1.02l-.77.78c-.18.18-.22.45-.09.68.74 1.3 1.77 2.38 3.03 3.19.23.15.52.14.72-.04l.82-.74c.29-.26.71-.34 1.06-.2l1.63.66c.4.16.67.53.68.96l.04 1.54c.01.5-.25.96-.69 1.2-.58.32-1.28.48-1.98.4-1.76-.2-3.71-1.32-5.55-3.16-1.83-1.83-2.96-3.78-3.16-5.54-.08-.71.08-1.41.4-1.99l.85-.94Z"
                          fill="#fff"
                        />
                      </svg>
                    </span>
                    <span>Ir a WhatsApp</span>
                    <span aria-hidden="true">↗</span>
                  </button>

                  <button type="button" className="info-catalog-button" onClick={() => { setInfoOpen(false); goHome(); }}>
                    <Home size={18} />
                    <span>Ir al inicio</span>
                    <span aria-hidden="true">↗</span>
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function CreditCardIcon() {
  return (
    <svg
      width="23"
      height="23"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <rect x="3" y="5" width="18" height="14" rx="2" />
      <path d="M3 10h18" />
      <path d="M7 15h3" />
    </svg>
  );
}

function QuantityInput({ value, max, onChange }) {
  const [draft, setDraft] = useState(String(value));

  useEffect(() => {
    setDraft(String(value));
  }, [value]);

  function handleChange(e) {
    const raw = e.target.value.replace(/\\D/g, "");
    setDraft(raw);

    if (raw !== "") {
      onChange(Number(raw));
    }
  }

  function handleBlur() {
    const number = Math.min(Math.max(1, Number(draft) || 1), max || 1);
    setDraft(String(number));
    onChange(number);
  }

  function handleKeyDown(e) {
    if (e.key === "Enter") {
      e.currentTarget.blur();
    }
  }

  return (
    <input
      className="qty-input"
      type="text"
      inputMode="numeric"
      value={draft}
      onChange={handleChange}
      onBlur={handleBlur}
      onKeyDown={handleKeyDown}
      aria-label="Cantidad"
    />
  );
}

function ProductCard({ product, onAdd, promo = false }) {
  const available = product.stock > 0;
  return (
    <article className="card">
      <div className="product-placeholder"><Package size={25} /></div>
      <div className="product-info">
        {promo && <em className="promo">PROMOCIÓN</em>}
        <b title={product.name}>{product.name}</b>
        <span>{product.family || product.category}</span>
        <strong>{money(product.price)}</strong>
        <em className={available ? "stock-ok" : "stock-no"}>{available ? "• En stock" : "• Sin stock"}</em>
        <button disabled={!available} onClick={() => onAdd(product)}><ShoppingCart size={13} /> Agregar</button>
      </div>
    </article>
  );
}

createRoot(document.getElementById("root")).render(<App />);
